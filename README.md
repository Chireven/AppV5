**Start-AppVApp.vbs**


When launching applications in a seamless environment (Citrix XenApp or Microsoft RemoteApp), it is difficult to publish App-V applications 
on a per-user basis.  In these environments, the application is often launched before the publishing configuration is complete causing a launch failure. 


These problems can be handled using this script to launch the application.  Instead of presenting the application to the user, you would present the customize
the script and publish the VBS instead.  The script will then do it's best to make sure the app is setup before it attempts to launch it.  It also tries it's 
best to prevent multiple publishing refreshes from happening at the same time.


To use the script:
1. Open Start-AppVApp.vbs in your favorite editor
2. Modify the following lines:


```vb
    appName       = "<your application name>"
    packageGUID   = "<your package GUID>"
    versionGUID   = "<your version GUID>
    exeToLaunch   = chr(34) & getAppvRoot & "\" & packageGUID & "\" & versionGUID & "\Root\VFS\myfolder\pathtoyour.exe"
    exeWorkingDir = getAppvRoot & "\" & packageGUID & "\" & versionGUID & "\Root\VFS\myfolder"
    exeParameters = ""      
``` 

=======
**appName**         : This is the display name.  Used in a friendly message to the user.  

**packageGUID**     : The App-V Package GUID.  Used to build the path to the App-V Package  

**versionGUID**     : The App-V Version GUID.  Used to build the path to the App-V package  

**exeToLaunch**     : Path to the EXE to launch when the script is run, typically points to the EXE in the App-V Package  

**exeWorkingDir**   : The path to launch the EXE from   

**exeParameters**   : Any paramaters that should be passed to the EXE when it's launched  

=======
  *appName*         : **This is the display name.  Used in a friendly message to the user.**
  *packageGUID*     : **The App-V Package GUID.  Used to build the path to the App-V Package**
  *versionGUID*     : **The App-V Version GUID.  Used to build the path to the App-V package**
  *exeToLaunch*     : **Path to the EXE to launch when the script is run, typically points to the EXE in the App-V Package**
  *exeWorkingDir*   : **The path to launch the EXE from** 
  *exeParameters*   : **Any paramaters that should be passed to the EXE when it's launched**




3. Save the new file as a new VBS file.

4. Configure your deployment software (XenApp/RemoteApp) to launched the vbscript instead of the EXE file


When the script is launched it will first look to see if the application is ready to run.  It does this by checking the publishing data in the users profile, as well as making sure
the App-V package is deployed to the machine.  If the app is ready to run, it will start it and exit the script.


If the App isn't ready to run, it will look to see if a Publishing Refresh is already happening.  If it can't detect one, it will 
initiate a publishing refresh and wait.


While the publishing refresh is happening, the variable midSyncLaunch comes in to play.  If it is enabled (default), the script 
will check at 1 second intervals to see if the app is ready.  If it becomes ready before the Publishing Refresh is complete,
it will launch the app and complete.  If midSyncLaunch is not enabled, the availability check will be done when the 
Publishing Refresh is complete.


Some Questions:

**Why vbscript?**
When publishing applications to a user, startup speed is very important.  When using powershell, there are delays introduced.  By using vbscript, we can 
reduce the amount of delays that are introduced by running a script before the application starts.


**Why hardcode the values?  Why not pass them on the command line?**
The average App-V path is long.  REALLY long.  By the time you add all the options that you want to a command line, you can quickly hit the limit of
most places where you would enter a command line (ie, a published application command line in a Citrix environment).  Hardcoding has it's own problems, 
but allows a much smaller command line.  
