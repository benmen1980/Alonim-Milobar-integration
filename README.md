# Alonim-Milobar-integration
A powershell script that post purchase order to MedaTech service
# How to execute the script from Priroity ?
:_PATH = 'C:\priority\bin.95\spix_api\script.ps1';

/* send command to powershell */
EXECUTE WINAPP 'c:\windows\system32',
/* for dubug mode
'cmd /c powershell -noexit',
*/
'cmd /c powershell',
'-command',:_PATH,
:_ORD,:_FILEPATH;


