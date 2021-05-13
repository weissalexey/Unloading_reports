'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.0
'
' NAME: strax.VBS
'
' AUTHOR: Alex Weiss,
' DATE  : 23.04.2007
'
' COMMENT: формирования отчетность для хранилища данных 
' создает лог C:\Work\log\OTchГГГГММДД.log
'==========================================================================

Dim localDateTime
Dim f
f = FormatDatetime(now(),2)
f = DateAdd("d", -1, f)
Dim t 
t = f
'WriteLog t
turdey=Date()
c_wkday = Weekday(Date())
WriteLog "*******Начало*******"
'''''''''''''''''''''''''''''''''''''''
If c_wkday = 1 Then WriteLog "Определяем день недели Воскресенье"
If c_wkday = 2 Then WriteLog "Определяем день недели Понедельник отчетность за Пятницу"
If c_wkday = 3 Then WriteLog "Определяем день недели Вторник отчетность за Понедельник"
If c_wkday = 4 Then WriteLog "Определяем день недели Среда отчетность за Вторник"
If c_wkday = 5 Then WriteLog "Определяем день недели Четверг отчетность за Среду"
If c_wkday = 6 Then WriteLog "Определяем день недели Пятница отчетность за Четверг"
If c_wkday = 7 Then WriteLog "Определяем день недели Суббота отчетность за Пятницу"
''''''''''''''''''''''''''''''''''''''''
'***************************************
If c_wkday = 2 Then 
t = DateAdd("d", -2, t)
WriteLog "Отчет за пятницу "& t
'***************************************
 c_year = Right(t,4)
 c_month = Right(Left(t,5),2)
 c_day = Left(t,2)
 If len(c_month)=1 Then c_month = "0" & c_month
 If len(c_day)=1 Then c_day = "0" & c_day
''''''''''''''''''''''''''''''''''''''''
'***************************************
sdp = c_year & c_month & c_day
	Fn = "V01" & right(t,1) & c_month & c_day & ".arj"
	Dim sql, result, result1
	Set con = CreateObject("ADODB.Connection")
	With con
		.Provider = "SQLOLEDB"
		.CommandTimeout = 0
		.Properties("Data Source") = "itan"
		.Properties("Integrated Security") = "SSPI"
		.Open
		.DefaultDatabase = "FIN"
	End With
	WriteLog "Смотрим Выгружался отчет за пятницу"
	'Dim sql,result
	sql = "select * from dbo.KHD_otch where date_pool =" & sdp & " and filName = '" & Fn & "'"
	WriteLog sql
	Set result1 =CreateObject ("ADODB.Recordset")
	result1.open sql,con
	If result1.EOF  Then
	WriteLog "Да отчет выгружался"
  '****************************************
	Else 
	t = DateAdd("d", 1, t)
	WriteLog "Выгружаем отчет за Субботу "& t
'***************************************
	End if
End If
'***************************************
	c_year = Right(t,4)
 	c_month = Right(Left(t,5),2)
 	c_day = Left(t,2)
 	If len(c_month)=1 Then c_month = "0" & c_month
 	If len(c_day)=1 Then c_day = "0" & c_day
	Dim fso, ts
	On Error Resume Next
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists("c:\odb.go\run\s1" & Right(c_year,2)& c_month & c_day &".021") Then
	WriteLog 
	make_dir "U:\BAL\XML\"& c_year
	make_dir "U:\BAL\XML\"& c_year & "\" & c_month 
	make_dir "U:\BAL\XML\"& c_year & "\" & c_month & "\"& c_day
	fso.MoveFile "c:\odb.go\run\s1" & Right(c_year,2)& c_month & c_day &".021","U:\BAL\XML\"& c_year & "\" & c_month & "\"&c_day & "\"& "s1" & right(t,2) & c_month & c_day & ".021"
	Bi = "U:"& Chr(13) & Chr(10)& _ 
		 "cd U:\BAL\XML\"& c_year & "\" & c_month& "\" & c_day & Chr(13) & Chr(10)& _
		 "arj a s1" & Right(c_year,2)& c_month & c_day &".arj *.021" 
		 	WriteBAT Bi
	Else
	WriteLog "НЕТ ФАЙЛА c:\odb.go\run\s1" & Right(c_year,2)& c_month & c_day &".021"
	end If
	' If fso.FileExists("c:\SBORNIK\input\VKLAD.121") Then
' 	make_dir "U:\BAL\sbornik\"& c_year
' 	make_dir "U:\BAL\sbornik\"& c_year & "\" & c_month 
' 	make_dir "U:\BAL\sbornik\"& c_year & "\" & c_month & "\"& c_day
' 	fso.MoveFile "c:\SBORNIK\input\VKLAD.121","U:\BAL\sbornik\"& c_year & "\" & c_month & "\"& c_day & "\"& "V01" & right(t,1) & c_month & c_day & ".121"
' 	Bi = "U:"& Chr(13) & Chr(10)& _ 
' 		 "cd U:\BAL\sbornik\"& c_year & "\" & c_month & Chr(13) & Chr(10)& _
' 		 "arj a V01" & right(t,1) & c_month & c_day & ".arj *.?21"
' 		 	WriteBAT Bi
' 	Else
' 	WriteLog "НЕТ ФАЙЛА c:\SBORNIK\input\VKLAD.121"
' 	end If
 '''''''''''''''''''''''''''''''''''
 ' If fso.FileExists("c:\SBORNIK\input\VKLAD.221") Then
' 	make_dir "U:\BAL\sbornik\"& c_year
' 	make_dir "U:\BAL\sbornik\"& c_year & "\" & c_month 
' 	make_dir "U:\BAL\sbornik\"& c_year & "\" & c_month & "\"& c_day
' 	fso.MoveFile "c:\SBORNIK\input\VKLAD.221","U:\BAL\sbornik\"& c_year & "\" & c_month & "\"& c_day & "\"&  "V01" & right(t,1) & c_month & c_day & ".221"
' 	Bi = "U:"& Chr(13) & Chr(10)& _ 
' 		 "cd U:\BAL\sbornik\"& c_year & "\" & c_month & "\" & c_day &  Chr(13) & Chr(10)& _
' 		 "arj a V01" & right(t,1) & c_month & c_day & ".arj *.?21"
' 		 	WriteBAT Bi
' 	Else
' 	WriteLog "НЕТ ФАЙЛА c:\SBORNIK\input\VKLAD.221"
' 	End If
	'********************************************************************
	sdp = c_year & c_month & c_day
	Fn = "s1" & Right(c_year,2)& c_month & c_day &".arj"
	Set con = CreateObject("ADODB.Connection")
	With con
		.Provider = "SQLOLEDB"
		.CommandTimeout = 0
		.Properties("Data Source") = "itan"
		.Properties("Integrated Security") = "SSPI"
		.Open
		.DefaultDatabase = "FIN"
	End With
	WriteLog "Смотрим на флаги в Таблице MS SQL  - <strax_otch>"
	sql = "select * from dbo.KHD_otch where date_pool =" & sdp & " and filName = '" & Fn & "'"
	WriteLog sql
	Set result =CreateObject ("ADODB.Recordset")
	result.open sql,con
	If result.EOF  Then
    sql ="insert into dbo.KHD_otch (Date_pool, FilName, Flag_poll) values ("& sdp &", '"& Fn & "',1 );"
    Set rez =CreateObject ("ADODB.Recordset")
    WriteLog sql
    rez.open sql,con
    WriteLog "*******Ставим флаги*******"
	End If
	WriteLog "*****Окончание отчета*****"
	
Sub WriteLog( param )'Создаем Log 
	Dim fso, ts
	Dim Dat , str
	Const ForAppend = 8
	c1_year = Year(Date())
	c1_month = Month(Date())
	c1_day = Day(Date())
	If len(c1_month)=1 Then c1_month = "0" & c1_month
	If len(c1_day)=1 Then c1_day = "0" & c1_day
	arc_name = c1_year & c1_month & c1_day 
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set ts = fso.OpenTextFile("C:\Work\log\OTch"& arc_name& ".log", ForAppend, True)
	'Dat = Hour(now())
    str = "[ " & FormatDateTime(now(),0) & " KHD_OTch ] "
    'dat = Minute(now())
	'str = str& Dat& " "
	str = str& param & Chr(13) & Chr(10)
	ts.Write (str)
	ts.Close
end Sub
Function make_dir(fldr)'Проверка и создание пути для архива
	Dim fso, ts, fs 
	Const ForWriting = 2
	On Error Resume Next
	Set fso = CreateObject("Scripting.FileSystemObject")
	if Not fso.FolderExists(fldr) Then
	   	Set fs =CreateObject("Scripting.FileSystemObject")
	  	Set ts = fs.CreateFolder(fldr)
	WriteLog "Каталог "& fldr & " Создан"& chr(13) & chr(10)
	Else
	WriteLog "Каталог "& fldr & " существует"& chr(13) & chr(10)
	end If
End Function 
Sub WriteBAT( param )'Создаем батник
	Dim fso, ts
	Const ForWrite  = 2
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set ts = fso.OpenTextFile("C:\SBORNIK\input\STRAX.bat", ForWrite, True)
	ts.Write (param)
	ts.Close
end Sub