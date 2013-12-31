'archive utility - Created on 24 -August-2012 by Karthikeyan.Sankarappan 

Dim objFSO 
strFilename = "emailable-report"
strExtn = ".html"
strFilename2 = "testng-failed"
strExtn2 = ".xml"

strSourceFolder = "c:\ebox\Karthik\2\MyProject\test-output" 
'strSourceFolder1 = Replace(strSourceFolder,"\","\\\\")

'Wscript.echo strSourceFolder1

strDestFolder = "c:\Documents and Settings\ksankarappan\Desktop\archive\" 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & _
        strComputer & "\root\cimv2")


Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
    ("Select * From __InstanceCreationEvent Within 5 Where " _
    & "Targetinstance Isa 'CIM_DirectoryContainsFile' and " _
    & "TargetInstance.GroupComponent= " _
    & "'Win32_Directory.Name=""c:\\\\ebox\\\\Karthik\\\\2\\\\MyProject\\\\test-output""'")


Do
    Set objLatestEvent = colMonitoredEvents.NextEvent


    temp = objLatestEvent.TargetInstance.PartComponent

	'Wscript.Echo temp

   temp = Replace(Mid(temp, InStr(temp, Chr(34)) + 1), "\\", "\") 

	'Wscript.Echo temp

     temp = Left(temp, Len(temp) - 1) 

	Wscript.Echo temp

	Wscript.Echo strSourceFolder & "\" & strFilename2 & strExtn


    if temp = strSourceFolder & "\" & strFilename & strExtn Then
	
		
	Set objFSO = CreateObject("Scripting.FileSystemObject") 

	Today = "_" + Replace(Date, "/", "_")  + "_" + Replace(FormatDateTime(Time, 3), ":","_")
 
	objFSO.MoveFile strSourceFolder & "\" & strFilename & strExtn, strDestFolder & strFilename & Today & strExtn

	'Wscript.echo "Emailable-Report.html is moved to the path : " &  strDestFolder 			

    End If
	
    if temp = strSourceFolder & "\" & strFilename2 & strExtn2 Then

		
	Set objFSO2 = CreateObject("Scripting.FileSystemObject") 

	Today = "_" + Replace(Date, "/", "_")  + "_" + Replace(FormatDateTime(Time, 3), ":","_")
 
	objFSO2.MoveFile strSourceFolder & "\" & strFilename2 & strExtn2, strDestFolder & strFilename2 & Today & strExtn2

	'Wscript.echo "testng-failed.xml is moved to the path : " &  strDestFolder 			


    End If
Loop
