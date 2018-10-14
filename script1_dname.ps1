# Script written by Mutu Adi-Marian fully licensed to Roy BenMenachem - ©2018
# Usage: powershell.exe -f "file_path_to_the_script" -o "ord value" -f "file path" -d "db name"
# You can found the log files in 'C:\priority\bin.95\simplyct'



param(
    [string]$o, # 'ord' parameter
    [string]$f, # 'filePath' parameter
    [string]$d # company name
);

$ord      = $o;
$filePath = $f;
$dname = $d;

# If the 'logs' directory doesn't exists we'll create it
if (!(Test-Path "C:\priority\bin.95\simplyct\logs\")) {
    mkdir "C:\priority\bin.95\simplyct\logs";
}

# Global Variables
$stamp = date -Format u;
$stamp = ($stamp.Substring(0,10)).Replace('-','');
$stamp = $stamp.Replace('-','');
$eLog  = "C:\priority\bin.95\simplyct\logs\$stamp-Error.log";
$rLog  = "C:\priority\bin.95\simplyct\logs\$stamp-Run.log";

# Functions

# Writes the current status to the console & log file
function Log {
    $stamp  = date -Format G;
    $stamp  = $stamp.Substring($stamp.Length-11,11)
    $strLog = "$stamp |> $args";

    Write-Host $strLog;
    Add-Content $rLog $strLog;
}

# Writes the occured error to the console & error log file
function ErrorLog {
    $stamp  = date -Format G;
    $stamp  = $stamp.Substring($stamp.Length-11,11);
    $strLog = "$stamp |> $args";

    # Writes to the main log file
    Add-Content $rLog "$stamp |> Fatal error occured. Operation aborted!";
    # Writes a more in-depth statement about the error to the error log file
    Add-Content $eLog $strLog;

   # Read-Host "$stamp |> $args`r`n`r`nPress ENTER to exit...";
    exit;
}

function ReadConfigFile {
    $configFilePath = "C:\priority\bin.95\simplyct\config.cfg";

    if (Test-Path $configFilePath) {
        $configFileContents = Get-Content $configFilePath;

        if ($configFileContents.Length -ne 7) {
            ErrorLog "Configuration file incomplete";
        }

        try {
            $global:serviceUrl = $configFileContents[0].Substring(11, $configFileContents[0].Length-11);
            $global:user       = $configFileContents[1].Substring(5,  $configFileContents[1].Length-5);
            $global:password   = $configFileContents[2].Substring(9,  $configFileContents[2].Length-9);
            $global:db_host    = $configFileContents[3].Substring(8,  $configFileContents[3].Length-8);
            $global:db_user    = $configFileContents[4].Substring(8,  $configFileContents[4].Length-8);
            $global:db_pass    = $configFileContents[5].Substring(8,  $configFileContents[5].Length-8);
            #$global:db_name    = $configFileContents[6].Substring(8,  $configFileContents[6].Length-8);
            $global:db_name    = $dname;
        } catch {
            ErrorLog 'Error while parsing the configuration file!';
        }
    } else{
        ErrorLog 'Configuration file missing!';
    }
}

function Execute-SOAPRequest { 
    try {
        Log "Sending SOAP Request to the service: $global:serviceUrl";

        $b64File                = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($filePath));
        $request                = [xml]"<?xml version=""1.0"" encoding=""utf-8""?><soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""><soap:Body><ReceiveOrderFiles xmlns=""https://www.miloubar.co.il""><inputOrderFile><UserName>$global:user</UserName><Password>$global:password</Password><FileContent>$b64File</FileContent><IsBase64>true</IsBase64></inputOrderFile></ReceiveOrderFiles></soap:Body></soap:Envelope>";

        $webRequest             = [System.Net.WebRequest]::Create($global:serviceUrl);
        $webRequest.Headers.Add("SOAPAction","`"https://www.miloubar.co.il/ReceiveOrderFiles`"");
        $webRequest.ContentType = "text/xml;charset=`"utf-8`"";
        $webRequest.Method      = "POST";
        $webRequest.Credentials = (New-Object System.Management.Automation.PSCredential($global:user, (ConvertTo-SecureString $global:password -AsPlainText -Force)));
        
        $requestStream          = $webRequest.GetRequestStream();
        $request.Save($requestStream);
        $requestStream.Close();
        
        Log "Request sent, getting the response..."; 

        $responseStream         = $webRequest.GetResponse().GetResponseStream();
        $soapReader             = [System.IO.StreamReader]($responseStream);
        $xmlResponse            = [xml]$soapReader.ReadToEnd();
        $responseStream.Close();

        Log "Response successfully received!";

        if ($xmlResponse.GetElementsByTagName("IsSuccess").'#text' -eq "true") {
            $global:service_Y = $xmlResponse.GetElementsByTagName("RefAU").'#text';

            return $true;
        }

        $global:service_error = $xmlResponse.GetElementsByTagName("ErrMsg").'#text';

        return $false;
    } catch {
        ErrorLog "SOAP Request failed: $_";
    }
}

# If one of the parameter are empty we'll tell to the user and ask to press the 'ENTER' key to exit
if ($ord -eq "") {
    ErrorLog "'ord' parameter is empty!";
}
if ($filePath -eq "") {
    ErrorLog "'filePath' parameter is empty!";
}
# If the file from the 'filePath' parameter doesn't exists
if (![System.IO.File]::Exists($filePath)) {
    ErrorLog "The '$filePath' file doesn't exists!";
}

# We read the entire configuration file
ReadConfigFile;

# Main
try {
    # Initializing the connection to the DB
    $connection                  = New-Object System.Data.SQLClient.SQLConnection;
    $connection.ConnectionString = "Server = $global:db_host; Database = $global:db_name; User ID = $global:db_user; Password = $global:db_pass;";

    $cmd            = New-Object System.Data.SQLClient.SQLCommand;
    $cmd.Connection = $connection;

    # Executing the SOAP Request
    if (Execute-SOAPRequest -eq $true) {
        # If the service response is 'TRUE'
        Log "Service Response: TRUE";

        $cmd.CommandText = "UPDATE PORDERS SET ELAR_WEBSERVICECONTR = '$global:service_Y' WHERE ORD = $ord";
    } else {
        # If the service response is 'FALSE'
        Log "Service Response: FALSE";
        Log "Service ErrorMsg: $global:service_error";

        $cmd.CommandText = "UPDATE PORDERS SET ELAR_WEBSERVICECONTR = 'ERR' WHERE ORD = $ord";
    }

    $connection.Open();
    $cmd.ExecuteReader();
    $connection.Close();

    Log "Database successfully updated!";
} catch {
    ErrorLog "Database: $_";
}
