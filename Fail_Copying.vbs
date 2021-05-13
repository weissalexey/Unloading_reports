
'***********************************************
' Создание объектов оболочки и файловой системы
Set oShell = CreateObject("wscript.shell")
Set oFSO = CreateObject("Scripting.Filesystemobject")
Set WSNetwork = CreateObject("WScript.Network")
LogFolder = "C:\work\log\" ' место расположения лог-файла
StartFolder = "\\Weiss-nb\upd\" ' откуда копируем
aEndFolder = array("C:\work\upd\") ' куда копируем

'***********************************************
' обнуление и описание счетчиков, используемых в скрипте
num_EndFolder = 0		'- общее число папок, места назначения для копируемых данных
num_EndFolder_0 = 0		'- не доступное кол-во папок из  num_EndFolder
num_files = 0			'- общее число обработанных файлов
num_files_copy = 0		'- из них скопировано с заменой на новую версию
err_files_copy = 0		'- из них не скопировано в результате ошибки при работе с num_files_copy
num_files_new = 0		'- из них скопировано новых файлов
err_files_new = 0		'- из них не скопировано в результате ошибки при работе с num_files_new
num_SubFolder = 0		'- обработано папок и подпапок
num_SubFolder_copy = 0	'- из них скопировано новых папок и подпапок
err_SubFolder = 0		'- из них не скопировано в результате ошибки при работе с num_SubFolder

'***********************************************
WriteLog "========== Запуск скрипта ==========" & vbCrLf

'***********************************************
Set oStartFolder = CreateObject("Scripting.FileSystemObject")
If oStartFolder.FolderExists ( StartFolder ) Then 
	'проверяю наличее каталога с файлами
	writelog StartFolder & " папка существует"
	Set oEndFolder = CreateObject("Scripting.FileSystemObject")
	' Цикл для перебора папок "куда копируем"
	For i=0 to UBound (aEndFolder)
		' Счетчик кол-ва папок для копируемых данных
		num_EndFolder=num_EndFolder+1
		' Проверяем доступность папки, в которую хотим произвести копирование
		If oEndFolder.FolderExists ( aEndFolder(i) ) Then 
			' Записываем результат в лог
			WriteLog "Папка " & "'" & aEndFolder(i) & "'" & " доступна для работы" & vbCrLf
			CopyFolder StartFolder,aEndFolder(i)
			' Счетчик доступных для копирования папок
			' удалил :)		
		else
			' Записываем результат в лог
		
			WriteLog "Папка " & "'" & aEndFolder(i) & "'" & " в настоящий момент не доступна. Работа с ней прекращена." & vbCrLf
			' Дублируем сообщение, выводом предупреждения на экран
			'WScript.Echo "Папка " & "'" & aEndFolder(i) & "'" & " в настоящий момент не доступна. Работа с ней прекращена."
			' Счетчик недоступных для копирования папок
			num_EndFolder_0=num_EndFolder_0+1
		End if
	Next

	WriteLog "========== Отчет о копировании ==========" & vbCrLf
	WriteLog "Было задано " & num_EndFolder & " папок, для копирования в них данных." 
	WriteLog "Часть папок оказалась не доступна - " & num_EndFolder_0 & " шт. (см.логи выше)"
	WriteLog "Скриптом было обработано - " & num_files & " файлов."
	WriteLog "Из них скопировано с заменой - " & num_files_copy & " шт., не скопировано в результате ошибки - " & err_files_copy & " шт."
	WriteLog "Из них скопировано новых файлов - " & num_files_new & " шт., не скопировано в результате ошибки - " & err_files_new & " шт."
	WriteLog "Скриптом было обработано - " & num_SubFolder & " подпапок."
	WriteLog "Из них скопировано новых подпапок - " & num_SubFolder_copy & " шт., не скопировано в результате ошибки - " & 	err_SubFolder & " шт."
	WriteLog "========== Выполнение скрипта завершено ==========" & vbCrLf
	'WScript.Echo "Выполнение скрипта завершено."

Sub CopyFolder(sCopyFolder,sEndCopyFolder)
	' Создание объекта Folder
    Set oFolder = oFSO.GetFolder(sCopyFolder)
	Set oEndCopyFolder = oFSO.GetFolder(sEndCopyFolder)
	' Получение коллекции файлов
    Set colFiles = oFolder.Files
    ' Обработка каждого файла из коллекции
    For each oFile in colFiles
		WriteLog "Дата создания копируемого файла:" & oFile.Name & vbTab & oFile.DateCreated
		' Счетчик числа проверяемых файлов
		num_files=num_files+1
		' Проверяем существует уже такой файл в папке, если его нет, то копируем. 
		' Если есть, то проверяем его актуальность и заменяем более новым, если он устарел.
		If oFSO.FileExists(oFSO.BuildPath(oEndCopyFolder, oFile.Name)) Then 
			' Записываем результат в лог
			WriteLog "Такой файл уже существует в папке " & oEndCopyFolder
			' Проверяем насколько это свежая копия файла, для этого сравниваем даты создания двух файлов
			WriteLog "Проверяем актуальность копии:"
			' Выгружаем полный путь к проверяемому файлу
			sFileEnd = oFSO.BuildPath(oEndCopyFolder, oFile.Name)
			' Создаем объект File, для работы с этим файлом
			Set oFileEnd = oFSO.GetFile(sFileEnd)
			' Сравниваем даты изменения файлов 
			If oFileEnd.DateLastModified < oFile.DateLastModified Then
				' Проверяемый файл оказался устаревшим, поэтому заменяем его более новым
				WriteLog "Копия файла устарела, заменяем его новым.     **********" & vbCrLf
				oFSO.CopyFile oFile, sEndCopyFolder & oFile.Name, True
				' Проверка на наличие ошибок
				if err.Number <> 0 then
					' Запись сообщения об ошибке в лог
					WriteLog "-----> Error # " & CStr(Err.Number) & " " & Err.Description
					' Очистка ошибки
					Err.Clear
					' Счетчик ошибок при замене файлов
					err_files_copy=err_files_copy+1
				else
					' Счетчик файлов, которые были заменены на новые
					num_files_copy=num_files_copy+1				
				End if
			else
				' В этом случае копия прошла проверку, просто продолжаем работу скрипта далее
				WriteLog "Копия актуальна. Продолжаем работу." & vbCrLf
			End if
		else
			' Записываем результат в лог
			WriteLog "Этот файл отсутствует в папке " & oEndCopyFolder & " Давайка его скопируем." & vbCrLf
			oFSO.CopyFile oFile, sEndCopyFolder & oFile.Name, True
			' Проверка на наличие ошибок
			if err.Number <> 0 then
				' Запись сообщения об ошибке в лог
				WriteLog "-----> Error # " & CStr(Err.Number) & " " & Err.Description
				' Очистка ошибки
				Err.Clear
				' Счетчик ошибок при копировании новых файлов
				err_files_new=err_files_copy+1
			else
				' Счетчик новых скопированных файлов
				num_files_new=num_files_copy+1				
			End if
		End if	
	Next
	' Проверяем все папки и подпапки
	WriteLog "Обрабатываем и копируем все подпапки из папки " & oEndCopyFolder & vbCrLf
	' Получение коллекции подпапок
    Set colSubFolders = oFolder.SubFolders
    ' Обработка каждой подпапки
    For Each oSubFolder In colSubFolders
		WriteLog "Проверяем подпапку " & oSubFolder
		' Счетчик обработанных папок и подпапок
		num_SubFolder=num_SubFolder+1	
		' Проверяем существует уже такая подпапка в папке, если ее нет, то копируем. 
		' Если есть, то переходим к проверке файлов в подпапке.
		If oFSO.FolderExists(oFSO.BuildPath(oEndCopyFolder, oFSO.GetBaseName(oSubFolder.Path))) Then
			' Записываем результат в лог
			WriteLog "Такая подпапка уже существует в папке " & oEndCopyFolder
			WriteLog "Проверяем все файлы в этой подпапке: "
			' Выгружаем полный путь к проверяемоой подпапке
			sSubFolderEnd = oFSO.BuildPath(oEndCopyFolder, oFSO.GetBaseName(oSubFolder.Path)) & "\" 
			' Производим рекурсивный вызов процедуры копирования файлов - программа вызывает сама себя
			CopyFolder oSubFolder, sSubFolderEnd			
			' oLogFile.Writeline
		else
			' Записываем результат в лог			
			WriteLog "Эта подпапка отсутствует в папке " & oEndCopyFolder & " Давайка ее скопируем." & vbCrLf
			oFSO.CopyFolder oSubFolder, sEndCopyFolder, True
			' Проверка на наличие ошибок
			if err.Number <> 0 then
				' Запись сообщения об ошибке в лог
				WriteLog "-----> Error # " & CStr(Err.Number) & " " & Err.Description
				' Очистка ошибки
				Err.Clear
				' Счетчик ошибок при копировании новых папок и подпапок
				err_SubFolder=err_SubFolder+1
			else
				' Счетчик новых скопированных папок и подпапок
				num_SubFolder_copy=num_SubFolder_copy+1			
			End if
		End if			
	Next
End Sub

Else
	WriteLog "Папка " & StartFolder & " не существует"
End if

Sub WriteLog(LogMessage)
Const ForAppending = 8
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C ="0" & C end if
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile("C:\work\log\" & A & B & C & ".log" , ForAppending, TRUE)
objLogFile.WriteLine("[" & Now() & "] " & LogMessage)
End Sub