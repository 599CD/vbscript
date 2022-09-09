' ---------------------------------
' Download a file from a website
'
' Author: Alex Hedley - 599CD.com
' Created: 3rd November 2012
' Updated: 9th September 2022
' ---------------------------------

' If you wanted the same thing everytime then you just need this one line of code.
' It will overwrite the file if it already exists
'download "http://599cd.s3-website-us-east-1.amazonaws.com/StudentFiles/Access/Access-2010-B1.zip", "D:\599CD\Sample\Access-2010-B1.zip"

'InputBox( prompt [, title] [, default] [, xpos] [, ypos] [, helpfile] [, context] )
' If you wish to download the file more than once prompt for an amount of times to loop
Dim TimesToDownload
TimesToDownload = InputBox ("Loop for ? times", "Loop For?", "3")

If TimesToDownload = "" Then
	Msgbox "You didn't chose a number.", vbCritical, "Quitting"
	WScript.Quit
End If
	
For i = 1 to TimesToDownload
	DownloadFile(i)
Next

' If you only wish to download it once you could comment out the above code and uncomment the below line.
' Sending an empty string "" will keep the filename as "Test.exe"
'downloadfile ""

Msgbox "All Complete", vbInformation, "Complete"

' ------------------------------
' Script to choose a FileName
' Calls the "download" function
' ------------------------------
Function DownloadFile(i)

	' Don't need a prompt for the Website as this is static, if you did you could do something like this:
	'Dim Website
	'InputBox( prompt [, title] [, default] [, xpos] [, ypos] [, helpfile] [, context] )
	'Website = InputBox ("Choose a Website Address" & vbNewLine & "and File to download:", "Website", "http://599cd.s3-website-us-east-1.amazonaws.com/StudentFiles/Access/Access-2010-B1.zip")
	Website = "http://599cd.s3-website-us-east-1.amazonaws.com/StudentFiles/Access/Access-2010-B1.zip"
	
	' Static FilePath
	Dim FilePath
	FilePath = "D:\599CD\Sample\"
	
	Dim FileName
	' Increment by the loop "i"
	'InputBox( prompt [, title] [, default] [, xpos] [, ypos] [, helpfile] [, context] )
	FileName = InputBox ("Path: " & FilePath & vbNewLine & vbNewLine & "Choose a FileName" _
		, "FileName and Location", "Test" & i & ".exe")

	' Join the FilePath to the FileName
	Dim File
	File = FilePath & FileName

	' Check that the User has chosen both
	'If IsNull(Website) Or IsNull(FileName) Then
	If Website = "" Or FileName = "" Then
		'MsgBox( prompt [, buttons] [, title] [, helpfile, context] )
		MsgBox "You haven't chosen a Website or FileName", vbExclamation, "Error"
		WScript.Quit
	Else
		' Call the "download" function with the website and file
		download Website, File
		' Inform the user it is complete
		'MsgBox( prompt [, buttons] [, title] [, helpfile, context] )
		Msgbox "Complete", vbInformation, "Complete"
	End If

End Function

' --------------------------
' Script to download a File
' --------------------------
function download(sFileURL, sLocation)
 
	'create xmlhttp object
	Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
 
	'get the remote file
	objXMLHTTP.open "GET", sFileURL, false
 
	'send the request
	objXMLHTTP.send()
 
	'wait until the data has downloaded successfully
	do until objXMLHTTP.Status = 200 :  wscript.sleep(1000) :  loop
 
	'if the data has downloaded sucessfully
	If objXMLHTTP.Status = 200 Then
 
		'create binary stream object
		Set objADOStream = CreateObject("ADODB.Stream")
		objADOStream.Open
 
		'adTypeBinary
		objADOStream.Type = 1
		objADOStream.Write objXMLHTTP.ResponseBody
 
		'Set the stream position to the start
		objADOStream.Position = 0    
 
		'create file system object to allow the script to check for an existing file
		Set objFSO = Createobject("Scripting.FileSystemObject")
 
	    'check if the file exists, if it exists then delete it
		If objFSO.Fileexists(sLocation) Then objFSO.DeleteFile sLocation
 
	    'destroy file system object
		Set objFSO = Nothing
 
	    'save the ado stream to a file
		objADOStream.SaveToFile sLocation
 
	    'close the ado stream
		objADOStream.Close
 
		'destroy the ado stream object
		Set objADOStream = Nothing
 
	'end object downloaded successfully
	End if
 
	'destroy xml http object
	Set objXMLHTTP = Nothing
 
End function

' Sample
'download "http://remote-location-of-file", "C:\name-of-file-and-extension"