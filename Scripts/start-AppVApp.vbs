' Init
  ' -----------------------------------------------------------------
  '   You need to customize these variables for Every Application
  ' -----------------------------------------------------------------
    appName       = "<your application name>"
    packageGUID   = "<your package GUID>"
    versionGUID   = "<your version GUID>
    exeToLaunch   = chr(34) & getAppvRoot & "\" & packageGUID & "\" & versionGUID & "\Root\VFS\myfolder\pathtoyour.exe"
    exeWorkingDir = getAppvRoot & "\" & packageGUID & "\" & versionGUID & "\Root\VFS\myfolder"
    exeParameters = ""        
  ' -----------------------------------------------------------------
  
  
  ' You can customize these variabes
  ' -----------------------------------------------------------------
    bDebug        = False
    midSyncLaunch = True
  
  ' Do not customize anything below this line.
  ' -----------------------------------------------------------------
    userSync    = vbTrue
    machineSync = vbTrue
    syncDone    = vbFalse
    doSync      = vbFalse
    Set oShell  = CreateObject("WScript.Shell")
    Set oFSO    = CreateObject("Scripting.FileSystemObject")
    Set oNet    = CreateObject("WScript.Network")
    Set ie      = CreateObject("InternetExplorer.Application")
  
  ' Customize IE
  ' -----------------------------------------------------------------    
    ie.AddressBar = 0    
    ie.MenuBar    = 0
    ie.Navigate   "About:Blank"
    ie.StatusBar  = 0
    ie.ToolBar    = 0
    
    ie.Width      = 400
    ie.Height     = 225
    
    Do While(ie.Busy)
      WScript.Sleep
    Loop
    
    
    ie.Document.Body.InnerHTML = generateWaitingMessage
    ie.Visible = 0    
  
 
' =============================================================================
' Check to see if a logon Sync is happening right now.  If it is, we will wait 
' for it to complete and we won't bother running another sync.
' =============================================================================
  syncWait = 0
  Alert "Checking Processes for Logon Publishing Refresh"
  If isLogonSyncRunning Then     
    Alert "  Already Running, we will not perform an additional Refresh"    
    userSync    = vbFalse
    machineSync = vbFalse
    ie.Visible  = 1
    
    Do Until isLogonSyncRunning = vbFalse
      Alert "    Waiting " & syncWait & " second(s) for sync"      
      WScript.Sleep 1000
      syncWait = syncWait + 1
      If midSyncLaunch Then
        If isPublishedToUser(packageGUID, versionGUID) Then Exit do
      End If
    Loop    
    syncDone = vbTrue
    On Error Resume Next
    Ie.visible = 0
    On Error Goto 0
    
  Else
    Alert "  No Refresh running.  One may be triggered if necessary"  
  End If

' =============================================================================
' Now that the sync is complete, we'll check to see if the user has permissions
' to run the application.  If they don't, we'll kick off a sync if one hasn't 
' been done before.
' =============================================================================
  Alert "Determine if the current user needs a Publishing Refresh"
  If isPublishedToUser(packageGUID, versionGUID) Then    
    Alert "  Publishing Information Found: User doesn't require a Publishing Refresh"
    userSync = vbFalse
  Else
    Alert "  Publishing Information Not Found: User Requires publishing Refresh"
    userSync  = vbTrue
  End if

' =============================================================================
' If the we didn't sync, but the package is missing from the machine, we will
' do a sync to bring the package down real quick.
' =============================================================================
  Alert "Determining if we need to sync to bring the package down to the machine"  
  If oFSO.FolderExists(getAppvRoot & "\" & packageGUID & "\" & versionGUID) Then
    Alert "  The package is on this machine, no need to sync."
    machineSync = vbFalse
  Else
    Alert "  The package is not on this machine.  A sync will be performed."
    machineSync = vbTrue
  End If
  
  
' =============================================================================
' If necessary, initate a publishing refresh
' =============================================================================  
  ' If we've already done a sync, there is no reason to do it again.
    If syncDone Then
      Alert "Publishing Refresh was initiated prior to this process.  It will be skipped."
    Else
      Alert "Publishing Refresh:"
      
      If userSync = vbTrue Then        
        Alert "  User is triggering Publishing Refresh"
        doSync = vbTrue
      End If
      
      If machineSync = vbTrue Then
        Alert "  Machine is triggering Publishing Refresh"
        doSync = vbTrue
      End If                                          
    End If
    
   If doSync Then     
     ie.Visible=1
     Alert "A Publishing Refresh is required."
     Alert "  Initiating a Publish Refresh with server 1"
     syncResults = syncServer     
     Alert "  Publishing Refresh Complete : [" & syncResults & "]"
     syncDone = vbTrue
     On Error Resume Next
     ie.Visible = 0
     On Error Goto 0
   Else
     Alert "A Publishing Refresh is not required"
   End If 


' =============================================================================
' At this point, we' should know for sure if a user has publsihing infromation 
' for the application.  We will check it, and then run the executable if the 
' user has permissions.  Do one last check to make sure ee can see the folder 
' and the executable, then run it.
' =============================================================================
  Alert "Checking Publishing Information"
  
  If isPublishedToUser(packageGUID, versionGUID) Then
    Alert "  Application is Published to User"
    Alert "    EXE               : [" & exeToLaunch & "]"
    Alert "    Arguments         : [" & exeParameters & "]"
    Alert "    Working Directory : [" & exeWorkingDir & "]"
    
    If exeWorkingDir <> "" Then
      If oFSO.FolderExists(exeWorkingDir) Then 
        Alert "    Working Directory : Exists"         
        oShell.CurrentDirectory = exeWorkingDir            
      Else
        Alert "    Working Directory : Does not Exist"         
        MsgBox "Could not find Working Directory to launch from", vbOKOnly & vbCritical, "Server: " & oNet.ComputerName
        Terminate
      End If
    End If
    
    Alert "Executing Program"
    oShell.Exec exeTolaunch & " " & exeParameters
    Terminate
    
  Else
    MsgBox "Publishing Information Unavailable.", vbOKOnly & vbCritical, "Server: " & oNet.ComputerName
  End If

' Terminate
  Terminate

' =============================================================================
' SUBS SUBS SUBS SUBS SUBS SUBS SUBS SUBS SUBS SUBS SUBS SUBS SUBS SUBS SUBS SU
' =============================================================================
Sub Terminate()
  ' Cleanup IE Object
    ie.Quit 
    Set ie=Nothing

  ' Terminate
    WScript.Quit
End Sub

' =============================================================================
' FUNCTIONS FUNCTIONS FUNCTIONS FUNCTIONS FUNCTIONS FUNCTIONS FUNCTIONS FUNCTIO
' =============================================================================

Function getAppvRoot()
' =============================================================================
' Purpose : Used to help locate the App-V Content on the local machine
' Returns : Base Path for App-V Content
' =============================================================================
  getAppvRoot   = "C:\ProgramData\App-V"
End Function

Function isPublishedToUser(packageGUID, versionGUID)
' =============================================================================
' Purpose : Used to check an application, by Pakcage and Version GUIDS, to
'         : determine if the application is published to the user.
' Returns : vbTrue/vbFalse
' =============================================================================
' Function to check and see if the publishing information exists
' in the users profile.
  Dim appData
  Dim catalogFolder
  
  Dim oFSO
  Dim oShell
        
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oShell = CreateObject("WScript.Shell")
  
  ' Get users appdata folder and clean it up
    appData = oShell.ExpandEnvironmentStrings("%appdata%")
    If Right(appData,1) <> "\" Then appData = appData & "\"
  
  ' Get users catalog folder and clean it up
    catalogFolder = appData & "Microsoft\AppV\Client\Catalog\Packages\"
    
  ' Determine the folder that should be avaliable if published to the user, and see if it exists.
    appFolder = catalogFolder & "{" & packageGUID & "}\{" & versionGUID & "}"
    If oFSO.FolderExists(appFolder) Then
      isPublishedToUser = vbTrue
      Alert "Publishing information Found"
    Else
      isPublishedToUser = vbFalse
      Alert "Publishing Information Not Found"
    End If     
End Function

Function generateWaitingMessage()
' =============================================================================
' Function : Generates the HTML message displayed in the Loading Dialog
' Returns  : string formatted in HTML
' =============================================================================
    generateWaitingMessage = ""
    generateWaitingMessage = generateWaitingMessage + "<html>"
    generateWaitingMessage = generateWaitingMessage + "  <head>"
    generateWaitingMessage = generateWaitingMessage + "    <title>" & appName & "</title>"
    generateWaitingMessage = generateWaitingMessage + "    <style type=" + Chr(34) + "text/css" + Chr(34) + ">"
    generateWaitingMessage = generateWaitingMessage + "       .loader {"
    generateWaitingMessage = generateWaitingMessage + "            margin: 10px auto;"
    generateWaitingMessage = generateWaitingMessage + "            font-size: 5px;"
    generateWaitingMessage = generateWaitingMessage + "            width: 1em;"
    generateWaitingMessage = generateWaitingMessage + "            height: 1em;"
    generateWaitingMessage = generateWaitingMessage + "            border-radius: 50%;"
    generateWaitingMessage = generateWaitingMessage + "            position: relative;"
    generateWaitingMessage = generateWaitingMessage + "            text-indent: -9999em;"
    generateWaitingMessage = generateWaitingMessage + "            -webkit-animation: load5 1.1s infinite ease;"
    generateWaitingMessage = generateWaitingMessage + "            -moz-animation: load5 1.1s infinite ease;"
    generateWaitingMessage = generateWaitingMessage + "            -o-animation: load5 1.1s infinite ease;"
    generateWaitingMessage = generateWaitingMessage + "            animation: load5 1.1s infinite ease;"
    generateWaitingMessage = generateWaitingMessage + "            -webkit-transform: translateZ(0);"
    generateWaitingMessage = generateWaitingMessage + "            -ms-transform: translateZ(0);"
    generateWaitingMessage = generateWaitingMessage + "            -moz-transform: translateZ(0);"
    generateWaitingMessage = generateWaitingMessage + "            -o-transform: translateZ(0);"
    generateWaitingMessage = generateWaitingMessage + "            transform: translateZ(0);"
    generateWaitingMessage = generateWaitingMessage + "        }"
    generateWaitingMessage = generateWaitingMessage + ""
    generateWaitingMessage = generateWaitingMessage + "        @-webkit-keyframes load5 {"
    generateWaitingMessage = generateWaitingMessage + "            0%, 100% { box-shadow: 0em -2.6em 0em 0em #00734a, 1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2), 2.5em 0em 0 0em rgba(0, 115, 75, 0.2), 1.75em 1.75em 0 0em rgba(0, 115, 75, 0.2), 0em 2.5em 0 0em rgba(0, 115, 75, 0.2), -1.8em 1.8em 0 0em rgba(0, 115, 75, 0.2), -2.6em 0em 0 0em rgba(0, 115, 75, 0.5), -1.8em -1.8em 0 0em rgba(0, 115, 75, 0.7);}"
    generateWaitingMessage = generateWaitingMessage + "            12.5%    { box-shadow: 0em -2.6em 0em 0em rgba(0, 115, 75, 0.7), 1.8em -1.8em 0 0em #00734a, 2.5em 0em 0 0em rgba(0, 115, 75, 0.2), 1.75em 1.75em 0 0em rgba(0, 115, 75, 0.2), 0em 2.5em 0 0em rgba(0, 115, 75, 0.2), -1.8em 1.8em 0 0em rgba(0, 115, 75, 0.2), -2.6em 0em 0 0em rgba(0, 115, 75, 0.2), -1.8em -1.8em 0 0em rgba(0, 115, 75, 0.5);}"
    generateWaitingMessage = generateWaitingMessage + "            25%      { box-shadow: 0em -2.6em 0em 0em rgba(0, 115, 75, 0.5), 1.8em -1.8em 0 0em rgba(0, 115, 75, 0.7), 2.5em 0em 0 0em #00734a, 1.75em 1.75em 0 0em rgba(0, 115, 75, 0.2), 0em 2.5em 0 0em rgba(0, 115, 75, 0.2), -1.8em 1.8em 0 0em rgba(0, 115, 75, 0.2), -2.6em 0em 0 0em rgba(0, 115, 75, 0.2), -1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2);}"
    generateWaitingMessage = generateWaitingMessage + "            37.5%    { box-shadow: 0em -2.6em 0em 0em rgba(0, 115, 75, 0.2), 1.8em -1.8em 0 0em rgba(0, 115, 75, 0.5), 2.5em 0em 0 0em rgba(0, 115, 75, 0.7), 1.75em 1.75em 0 0em rgba(0, 115, 75, 0.2), 0em 2.5em 0 0em rgba(0, 115, 75, 0.2), -1.8em 1.8em 0 0em rgba(0, 115, 75, 0.2), -2.6em 0em 0 0em rgba(0, 115, 75, 0.2), -1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2);}"
    generateWaitingMessage = generateWaitingMessage + "            50%      { box-shadow: 0em -2.6em 0em 0em rgba(0, 115, 75, 0.2), 1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2), 2.5em 0em 0 0em rgba(0, 115, 75, 0.5), 1.75em 1.75em 0 0em rgba(0, 115, 75, 0.7), 0em 2.5em 0 0em #00734a, -1.8em 1.8em 0 0em rgba(0, 115, 75, 0.2), -2.6em 0em 0 0em rgba(0, 115, 75, 0.2), -1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2);}"
    generateWaitingMessage = generateWaitingMessage + "            62.5%    { box-shadow: 0em -2.6em 0em 0em rgba(0, 115, 75, 0.2), 1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2), 2.5em 0em 0 0em rgba(0, 115, 75, 0.2), 1.75em 1.75em 0 0em rgba(0, 115, 75, 0.5), 0em 2.5em 0 0em rgba(0, 115, 75, 0.7), -1.8em 1.8em 0 0em #00734a, -2.6em 0em 0 0em rgba(0, 115, 75, 0.2), -1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2);}"
    generateWaitingMessage = generateWaitingMessage + "            75% 	 { box-shadow: 0em -2.6em 0em 0em rgba(0, 115, 75, 0.2), 1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2), 2.5em 0em 0 0em rgba(0, 115, 75, 0.2), 1.75em 1.75em 0 0em rgba(0, 115, 75, 0.2), 0em 2.5em 0 0em rgba(0, 115, 75, 0.5), -1.8em 1.8em 0 0em rgba(0, 115, 75, 0.7), -2.6em 0em 0 0em #00734a, -1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2);}"
    generateWaitingMessage = generateWaitingMessage + "            87.5%    { box-shadow: 0em -2.6em 0em 0em rgba(0, 115, 75, 0.2), 1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2), 2.5em 0em 0 0em rgba(0, 115, 75, 0.2), 1.75em 1.75em 0 0em rgba(0, 115, 75, 0.2), 0em 2.5em 0 0em rgba(0, 115, 75, 0.2), -1.8em 1.8em 0 0em rgba(0, 115, 75, 0.5), -2.6em 0em 0 0em rgba(0, 115, 75, 0.7), -1.8em -1.8em 0 0em #00734a;}"
    generateWaitingMessage = generateWaitingMessage + "        }"
    generateWaitingMessage = generateWaitingMessage + ""
    generateWaitingMessage = generateWaitingMessage + "        @keyframes load5 {"
    generateWaitingMessage = generateWaitingMessage + "            0%, 100% { box-shadow: 0em -2.6em 0em 0em #00734a, 1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2), 2.5em 0em 0 0em rgba(0, 115, 75, 0.2), 1.75em 1.75em 0 0em rgba(0, 115, 75, 0.2), 0em 2.5em 0 0em rgba(0, 115, 75, 0.2), -1.8em 1.8em 0 0em rgba(0, 115, 75, 0.2), -2.6em 0em 0 0em rgba(0, 115, 75, 0.5), -1.8em -1.8em 0 0em rgba(0, 115, 75, 0.7);}"
    generateWaitingMessage = generateWaitingMessage + "            12.5%    { box-shadow: 0em -2.6em 0em 0em rgba(0, 115, 75, 0.7), 1.8em -1.8em 0 0em #00734a, 2.5em 0em 0 0em rgba(0, 115, 75, 0.2), 1.75em 1.75em 0 0em rgba(0, 115, 75, 0.2), 0em 2.5em 0 0em rgba(0, 115, 75, 0.2), -1.8em 1.8em 0 0em rgba(0, 115, 75, 0.2), -2.6em 0em 0 0em rgba(0, 115, 75, 0.2), -1.8em -1.8em 0 0em rgba(0, 115, 75, 0.5);}"
    generateWaitingMessage = generateWaitingMessage + "            25%      { box-shadow: 0em -2.6em 0em 0em rgba(0, 115, 75, 0.5), 1.8em -1.8em 0 0em rgba(0, 115, 75, 0.7), 2.5em 0em 0 0em #00734a, 1.75em 1.75em 0 0em rgba(0, 115, 75, 0.2), 0em 2.5em 0 0em rgba(0, 115, 75, 0.2), -1.8em 1.8em 0 0em rgba(0, 115, 75, 0.2), -2.6em 0em 0 0em rgba(0, 115, 75, 0.2), -1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2);}"
    generateWaitingMessage = generateWaitingMessage + "            37.5%    { box-shadow: 0em -2.6em 0em 0em rgba(0, 115, 75, 0.2), 1.8em -1.8em 0 0em rgba(0, 115, 75, 0.5), 2.5em 0em 0 0em rgba(0, 115, 75, 0.7), 1.75em 1.75em 0 0em rgba(0, 115, 75, 0.2), 0em 2.5em 0 0em rgba(0, 115, 75, 0.2), -1.8em 1.8em 0 0em rgba(0, 115, 75, 0.2), -2.6em 0em 0 0em rgba(0, 115, 75, 0.2), -1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2);}"
    generateWaitingMessage = generateWaitingMessage + "            50%      { box-shadow: 0em -2.6em 0em 0em rgba(0, 115, 75, 0.2), 1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2), 2.5em 0em 0 0em rgba(0, 115, 75, 0.5), 1.75em 1.75em 0 0em rgba(0, 115, 75, 0.7), 0em 2.5em 0 0em #00734a, -1.8em 1.8em 0 0em rgba(0, 115, 75, 0.2), -2.6em 0em 0 0em rgba(0, 115, 75, 0.2), -1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2);}"
    generateWaitingMessage = generateWaitingMessage + "            62.5%    { box-shadow: 0em -2.6em 0em 0em rgba(0, 115, 75, 0.2), 1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2), 2.5em 0em 0 0em rgba(0, 115, 75, 0.2), 1.75em 1.75em 0 0em rgba(0, 115, 75, 0.5), 0em 2.5em 0 0em rgba(0, 115, 75, 0.7), -1.8em 1.8em 0 0em #00734a, -2.6em 0em 0 0em rgba(0, 115, 75, 0.2), -1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2); }"            
    generateWaitingMessage = generateWaitingMessage + "            75%      { box-shadow: 0em -2.6em 0em 0em rgba(0, 115, 75, 0.2), 1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2), 2.5em 0em 0 0em rgba(0, 115, 75, 0.2), 1.75em 1.75em 0 0em rgba(0, 115, 75, 0.2), 0em 2.5em 0 0em rgba(0, 115, 75, 0.5), -1.8em 1.8em 0 0em rgba(0, 115, 75, 0.7), -2.6em 0em 0 0em #00734a, -1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2);}"          
    generateWaitingMessage = generateWaitingMessage + "            87.5%    { box-shadow: 0em -2.6em 0em 0em rgba(0, 115, 75, 0.2), 1.8em -1.8em 0 0em rgba(0, 115, 75, 0.2), 2.5em 0em 0 0em rgba(0, 115, 75, 0.2), 1.75em 1.75em 0 0em rgba(0, 115, 75, 0.2), 0em 2.5em 0 0em rgba(0, 115, 75, 0.2), -1.8em 1.8em 0 0em rgba(0, 115, 75, 0.5), -2.6em 0em 0 0em rgba(0, 115, 75, 0.7), -1.8em -1.8em 0 0em #00734a; }"
    generateWaitingMessage = generateWaitingMessage + "        }"
    generateWaitingMessage = generateWaitingMessage + ""
    generateWaitingMessage = generateWaitingMessage + "	    tr.allBorders td"
    generateWaitingMessage = generateWaitingMessage + "		  { border-bottom:2px solid black;"
    generateWaitingMessage = generateWaitingMessage + "		    border-top:2px solid black;"
    generateWaitingMessage = generateWaitingMessage + "			border-left:2px solid black;"
    generateWaitingMessage = generateWaitingMessage + "			border-right:2px solid black;"
    generateWaitingMessage = generateWaitingMessage + "		}"
    generateWaitingMessage = generateWaitingMessage + ""
    generateWaitingMessage = generateWaitingMessage + "		body {"
    generateWaitingMessage = generateWaitingMessage + "			font-family:Arial, Helvetica, sans-serif;"
    generateWaitingMessage = generateWaitingMessage + "			background-color:white;"
    generateWaitingMessage = generateWaitingMessage + "    </style>"
    generateWaitingMessage = generateWaitingMessage + "  </head>"
    generateWaitingMessage = generateWaitingMessage + "<body>"
    generateWaitingMessage = generateWaitingMessage + "	<table align=" & chr(34)  & "center" & chr(34)  & " cellpadding=" & chr(34)  & "5" & chr(34)  & " border=0>"
    generateWaitingMessage = generateWaitingMessage + "		<tbody>"
    generateWaitingMessage = generateWaitingMessage + "			<tr class=" & chr(34)  & "allBorders" & chr(34)  & ">"
    generateWaitingMessage = generateWaitingMessage + "				<td align=" & chr(34)  & "center" & chr(34)  & " style=" & chr(34)  & "font-size: 20px" & chr(34)  & " bgcolor=" & chr(34)  & "#339966" & chr(34)  & " colspan=" & chr(34)  & "100%" & chr(34)  & " >"
    generateWaitingMessage = generateWaitingMessage + "					<b>Please Wait</b>"
    generateWaitingMessage = generateWaitingMessage + "				</td>"
    generateWaitingMessage = generateWaitingMessage + "			</tr>"
    generateWaitingMessage = generateWaitingMessage + "			<tr>"
    generateWaitingMessage = generateWaitingMessage + "				<td align=" & chr(34)  & "center" & chr(34)  & " width = " & chr(34)  & "50px" & chr(34)  & "><div class=" & chr(34)  & "loader" & chr(34)  & ">Loading...</div></td>"
    generateWaitingMessage = generateWaitingMessage + ""
    generateWaitingMessage = generateWaitingMessage + "				<td align=" & chr(34)  & "left" & chr(34)  & ">"
    generateWaitingMessage = generateWaitingMessage + "				  <div><i>While we get your application ready to launch...<i><br><br></div>"
    generateWaitingMessage = generateWaitingMessage + "				  <div>"
    generateWaitingMessage = generateWaitingMessage + "				    <table border=0 cellpadding=4>"
    generateWaitingMessage = generateWaitingMessage + "				       <tr>"
    generateWaitingMessage = generateWaitingMessage + "				          <td><b>Application</b></td>"
    generateWaitingMessage = generateWaitingMessage + "				          <td><i>" & appName & "</i></td></tr></div>"
    generateWaitingMessage = generateWaitingMessage + "				       </tr>"
    generateWaitingMessage = generateWaitingMessage + "				       <tr>"
    generateWaitingMessage = generateWaitingMessage + "				         <td><B>Server</td>"
    generateWaitingMessage = generateWaitingMessage + "				         <td>" & onet.ComputerName & "</td>"    
    generateWaitingMessage = generateWaitingMessage + "				       </tr>"
    generateWaitingMessage = generateWaitingMessage + "				   </table>"    
    generateWaitingMessage = generateWaitingMessage + "				</td>"
    generateWaitingMessage = generateWaitingMessage + "			</tr>"
    generateWaitingMessage = generateWaitingMessage + ""
    generateWaitingMessage = generateWaitingMessage + "		</tbody>"
    generateWaitingMessage = generateWaitingMessage + "	</table>"
    generateWaitingMessage = generateWaitingMessage + ""
    generateWaitingMessage = generateWaitingMessage + "</body>"    
    generateWaitingMessage = generateWaitingMessage + "</html>"
            
End Function


Function syncServer()
' =============================================================================
' Function : Initiates a Sync with the App-V Publishing Server
' Returns  : Exit code from the SyncAppvPublishingServer.exe program
' =============================================================================
  Dim syncEXE
  Dim syncParamaters
  Dim syncWorkingDirectory
  Dim syncCommand
  
  Dim oFSO
  Dim oShell
  Dim oExec
  
  syncEXE              = "SyncAppvPublishingServer.exe"
  syncParameters       = "1 -NetworkCostAware"
  syncWorkingDirectory = "C:\Program Files\Microsoft Application Virtualization\Client\"
    
  Set oFSO   = CreateObject("Scripting.FileSystemObject")
  Set oShell = CreateObject("WScript.Shell")
  
  If Right(syncWorkingDirectory, 1) <> "\" Then syncWorkingDirectory = syncWorkingDirectory & "\"
  syncCommand = syncWorkingDirectory & syncEXE 
  
  If Not oFSO.FileExists(syncCommand) Then
    Alert "[" & syncCommand & "] not found"
    syncServer = vbFalse
    Exit Function    
  End If
  
  oShell.CurrentDirectory = syncWorkingDirectory
  Alert "  Initiating Sync"
  Alert "    " & syncEXE & " " & syncParameters
  Set oExec	= oShell.Exec(syncEXE & " " & syncParameters)
  
  Do While oExec.Status = 0
    WScript.Sleep 1000
    Alert "    Waiting " & syncWait & " second(s) for sync"
    syncWait = syncWait + 1
    If midSyncLaunch Then
      If isPublishedToUser(packageGUID, versionGUID) Then 
        syncServer = -1
        Exit Do
      End If
    End If
  Loop
  Alert "  Sync Complete"
 

   syncServer = oExec.ExitCode
   Exit Function  
End Function

Function isLogonSyncRunning()
' =============================================================================
' Purpose : Determines if a current instances of a logon publishing refresh
'         : is currently running for the current user.
' Returns : vbTrue/vbFalse
' =============================================================================
  Dim syncProcess
  Dim query 
  Dim oWMI
  Dim oNet
  
  syncProcess    = "SyncAppvPublishingServer.exe"
  query          ="SELECT * FROM Win32_Process WHERE NAME = '" & syncProcess & "'"
  
  Set oWMI       = GetObject ("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  Set oProcesses = oWMI.ExecQuery(query)
  Set oNet       = CreateObject("WScript.Network")
  
  If oProcesses.count > 0 Then      
   For Each process In oProcesses
     On Error Resume Next
     Owner = vbNull           
     O = Process.GetOwner(Owner)    
     On Error Goto 0     
     If LCase(Owner) = LCase(oNet.username) Then
       isLogonSyncRunning = vbTrue       
       Exit Function
     End If       
   Next
  End If
  ' Username doesn't match, or we haven't found this user to be the owner.
  isLogonSyncRunning = vbFalse       
End Function

Function Alert(message)
' =============================================================================
' Purpose : Logs an Alert
' Returns : Nothing
' =============================================================================
  If bDebug Then WScript.Echo Message
End Function
