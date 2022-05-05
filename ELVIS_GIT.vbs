PatchPDF ="c:\Users\aw\Desktop\winspied\ELVIS\"
FolderTo ="c:\Users\aw\Desktop\winspied\ELVIS_BK\"
'PatchPDF ="E:\Schnittstellen\Elvis\OUT\POD\"
sql = "select * from V_POD_ELVIS_941"
WriteLog "Start"
Set con = CreateObject("ADODB.Connection")
	
	With con
		.Provider = "SQLOLEDB"
		.Properties("Data Source") = "(local)"
		.Properties("Integrated Security") = "SSPI"
		.Open
		.DefaultDatabase = "Fin"
	End With
   
   Set result = con.Execute(sql)
   

    If Not result.EOF  Then

	  result.MoveFirst
	  While Not result.EOF
	  	dim LN, naim, FE, DOCDAT
	  	  SWort = result.Fields("SWort").Value 
	  	  ZusText_E = result.Fields("ZusText_E").Value
          ZusText_E = replace (ZusText_E, "/", "-")
          DOCDAT = result.Fields("DocumentData").Value
          'MsgBox DOCDAT
        SaveBinaryData  PatchPDF &"POD_" & SWort & "_" & ZusText_E & "_1.PDF", DOCDAT
	    WriteLog PatchPDF &"POD_" & SWort & "_" & ZusText_E & "_1.PDF"
    
	   result.movenext
	  wend
      FTPUpload(PatchPDF) 
      BackUpFile PatchPDF, FolderTo
	end if
writelog " End"

Function SaveBinaryData(FileName, ByteArray)
  Const adTypeBinary = 1
  Const adSaveCreateOverWrite = 2
  
  'Create Stream object
  Dim BinaryStream
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To save binary data.
  BinaryStream.Type = adTypeBinary
  
  'Open the stream And write binary data To the object
  BinaryStream.Open
  BinaryStream.Write ByteArray
  
  'Save binary data To disk
  
  BinaryStream.SaveToFile FileName, adSaveCreateOverWrite
End Function

Sub WriteLog(LogMessage)
Const ForAppending = 8
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C ="0" & C end if
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile("" & A & B & C & ".log" , ForAppending, TRUE)
objLogFile.WriteLine("[" & Now() & "_ELVIS] " & LogMessage)
End Sub

'

Sub FTPUpload(path)
Set oShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next

Const copyType = 16

'FTP Wait Time in ms
waitTime = 80000

FTPUser = ""
FTPPass = ""
FTPHost = ""
FTPDir = ""

strFTP = "ftp://" & FTPUser & ":" & FTPPass & "@" & FTPHost & FTPDir
Set objFTP = oShell.NameSpace(strFTP)


'Upload single file       
If objFSO.FileExists(path) Then

Set objFile = objFSO.getFile(path)
strParent = objFile.ParentFolder
Set objFolder = oShell.NameSpace(strParent)

Set objItem = objFolder.ParseName(objFile.Name)

writelog "Uploading file " & objItem.Name & " to " & strFTP
 objFTP.CopyHere objItem, copyType


End If


'Upload all files in folder
If objFSO.FolderExists(path) Then

'Entire folder
Set objFolder = oShell.NameSpace(path)

writelog "Uploading folder " & path & " to " & strFTP
objFTP.CopyHere objFolder.Items, copyType

End If


If Err.Number <> 0 Then
WriteLog "Error: " & Err.Description
End If

'Wait for upload
Wscript.Sleep waitTime

End Sub


sub BackUpFile (PatchPDF,FolderTo )
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C ="0" & C end if

Set FSO=CreateObject("Scripting.FileSystemObject")
Set fldr= FSO.GetFolder(PatchPDF)
Set Collec_Files= fldr.Files
set fso1=CreateObject ("Scripting.FileSystemObject") 
For Each File in Collec_Files
    If Collec_Files.count < 3 then
      Writelog "Need more files"
    Else
      set new_folder=fso1.CreateFolder(FolderTo & A & B & C & "\")
      FSO.MoveFile PatchPDF & "*", FolderTo & A & B & C & "\" 
    End If
Next


End sub

