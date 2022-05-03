pathcXML ="c:\Users\aw\Desktop\winspied\ELVIS\"
PatchJpg ="c:\Users\aw\Desktop\winspied\ELVIS\"
Y = year(Now)
M=Month(Now)
if len(M) < 2  then M="0" & M end if
D = day (Now)
if len(D)<2 then D ="0" & D end if
'YMD = Y & M & D & "_"


 	Set con = CreateObject("ADODB.Connection")
	With con
		.Provider = "SQLOLEDB"
		.Properties("Data Source") = "(local)"
		.Properties("Integrated Security") = "SSPI"
		.Open
		.DefaultDatabase = "Fin"
	End With

   sql = "select * from V_POD_ELVIS_941"
   Set result = con.Execute(sql)
   

    If Not result.EOF  Then

	  result.MoveFirst
	  While Not result.EOF
	  	dim LN, naim, FE, DOCDAT
	  	  SWort = result.Fields("SWort").Value 
	  	  ZusText_E = result.Fields("ZusText_E").Value
          DOCDAT = result.Fields("DocumentData").Value
          'MsgBox DOCDAT
        SaveBinaryData  SWort & "_" & ZusText_E & ".PDF", DOCDAT
	
	   result.movenext
	  wend
	end if

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
