 
      	Set con = CreateObject("ADODB.Connection")
	With con
		.Provider = "SQLOLEDB"
		.Properties("Data Source") = "(local)"
		.Properties("Integrated Security") = "SSPI"
		.Open
		.DefaultDatabase = "Fin"
	End With
	
	Dim fso, ts, filespec
	filespec = "C:\Documents and Settings\familia\Рабочий стол\qwerty\Index1.htm"
	Const ForWriting = 2
	Set fso = CreateObject("Scripting.FileSystemObject")
	If (fso.FileExists(filespec)) Then fso.DeleteFile(filespec)
 	Set ts = fso.OpenTextFile(filespec, ForWriting, True) 
	dim sql,rs 
	sql = "Select * From fin..stores"
	Set rs = con.Execute(sql)
	ts.WriteLine "<HEAD> "
	ts.WriteLine ""
	ts.WriteLine "<H4>Таблица изменений</H4> "
	ts.WriteLine ""
	ts.WriteLine "<BODY> "
	ts.WriteLine ""
	ts.WriteLine "<TABLE CellPadding=1 CellSpacing=1 Cols=5> "
	ts.WriteLine "<TBODY> "
	ts.WriteLine "<TR VALIGN=top ALIGN=left> "
	ts.WriteLine "<TH>КОД</TH> "
	ts.WriteLine "<TH>ИМЯ</TH> "
	ts.WriteLine "<TH>АДРЕС</TH> "
	ts.WriteLine "<TH>ГОРОД</TH> "
	ts.WriteLine "<TH>ШТАТ</TH> "
	ts.WriteLine "<TH>Zip КОД</TH> "
	ts.WriteLine "</TR> "
	ts.WriteLine "<TBODY> "
	
	
	If Not rs.EOF  Then
	  rs.MoveFirst
	  While Not rs.EOF
	  	dim cod,naim
	  	  

a =rs.Fields("stor_id").Value 
b =rs.Fields("stor_name").Value 
c =rs.Fields("stor_address").Value 
d =rs.Fields("city").Value 
e =rs.Fields("state").Value 
f =rs.Fields("zip").Value
	  	  
ts.WriteLine "<TR VALIGN=top ALIGN=left> "	  	  
ts.WriteLine "<TD>" & a & "</TD>" 
ts.WriteLine "<TD>" & b & "</TD>" 
ts.WriteLine "<TD>" & c & "</TD>" 
ts.WriteLine "<TD>" & d & "</TD>" 
ts.WriteLine "<TD>" & e & "</TD>" 
ts.WriteLine "<TD>" & f & "</TD>" 
ts.WriteLine "</TR> "	  	  
	  	  
	  	  
  	  rs.movenext
	  wend
	end if

ts.WriteLine "</TABLE> "
ts.WriteLine "</BODY> "
ts.WriteLine "</HTML> "
	
	ts.Close
