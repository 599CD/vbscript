' ---------------------------------------------------------------------
' Rename a file and move it
' Rename it to datetime and move to archive.
'
' Created by Alex Hedley
' Date Created: 28/01/2013
' Date Updated: ##/##/####
'
' Required Refs: 
' ---------------------------------------------------------------------
Dim fso 
Set fso = CreateObject("Scripting.FileSystemObject") 
  
Dim filename
'filename = "testfile.txt"
filename = "Report.csv"
  
Dim path
path = "ADD YOUR PATH HERE"
  
Dim demofile
Set demofile = fso.GetFile(path & filename)
createdate = demofile.DateCreated

' Format DateTime from dd/mm/yyyy hh:nn:ss => yyyymmdd_hhnn
'createdate = year(createdate) & month(createdate) & _
'   day(createdate) & "_" & hour(createdate) & minute(createdate)

'Build up the Date
Dim yearNum, monthNum, dayNum, hourNum, minNum
yearNum = year(createdate)
monthNum = month(createdate)
dayNum = day(createdate)
hourNum = hour(createdate)
minNum = minute(createdate)

' 0s are going, add them back in.
If len(monthNum) < 2 Then monthNum = "0" & monthNum
If len(dayNum) < 2 Then dayNum = "0" & dayNum
If len(hourNum) < 2 Then hourNum = "0" & hourNum
If len(minNum) < 2 Then minNum = "0" & minNum
' Join it up
createdate = yearNum & monthNum & dayNum & "_" & hourNum & minNum

' Add the CreateDate and move to the archive
fso.MoveFile filename, path & "archive\Report " & createdate & ".csv"

Msgbox "Done"