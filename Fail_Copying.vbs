
'***********************************************
' �������� �������� �������� � �������� �������
Set oShell = CreateObject("wscript.shell")
Set oFSO = CreateObject("Scripting.Filesystemobject")
Set WSNetwork = CreateObject("WScript.Network")
LogFolder = "C:\work\log\" ' ����� ������������ ���-�����
StartFolder = "\\Weiss-nb\upd\" ' ������ ��������
aEndFolder = array("C:\work\upd\") ' ���� ��������

'***********************************************
' ��������� � �������� ���������, ������������ � �������
num_EndFolder = 0		'- ����� ����� �����, ����� ���������� ��� ���������� ������
num_EndFolder_0 = 0		'- �� ��������� ���-�� ����� ��  num_EndFolder
num_files = 0			'- ����� ����� ������������ ������
num_files_copy = 0		'- �� ��� ����������� � ������� �� ����� ������
err_files_copy = 0		'- �� ��� �� ����������� � ���������� ������ ��� ������ � num_files_copy
num_files_new = 0		'- �� ��� ����������� ����� ������
err_files_new = 0		'- �� ��� �� ����������� � ���������� ������ ��� ������ � num_files_new
num_SubFolder = 0		'- ���������� ����� � ��������
num_SubFolder_copy = 0	'- �� ��� ����������� ����� ����� � ��������
err_SubFolder = 0		'- �� ��� �� ����������� � ���������� ������ ��� ������ � num_SubFolder

'***********************************************
WriteLog "========== ������ ������� ==========" & vbCrLf

'***********************************************
Set oStartFolder = CreateObject("Scripting.FileSystemObject")
If oStartFolder.FolderExists ( StartFolder ) Then 
	'�������� ������� �������� � �������
	writelog StartFolder & " ����� ����������"
	Set oEndFolder = CreateObject("Scripting.FileSystemObject")
	' ���� ��� �������� ����� "���� ��������"
	For i=0 to UBound (aEndFolder)
		' ������� ���-�� ����� ��� ���������� ������
		num_EndFolder=num_EndFolder+1
		' ��������� ����������� �����, � ������� ����� ���������� �����������
		If oEndFolder.FolderExists ( aEndFolder(i) ) Then 
			' ���������� ��������� � ���
			WriteLog "����� " & "'" & aEndFolder(i) & "'" & " �������� ��� ������" & vbCrLf
			CopyFolder StartFolder,aEndFolder(i)
			' ������� ��������� ��� ����������� �����
			' ������ :)		
		else
			' ���������� ��������� � ���
		
			WriteLog "����� " & "'" & aEndFolder(i) & "'" & " � ��������� ������ �� ��������. ������ � ��� ����������." & vbCrLf
			' ��������� ���������, ������� �������������� �� �����
			'WScript.Echo "����� " & "'" & aEndFolder(i) & "'" & " � ��������� ������ �� ��������. ������ � ��� ����������."
			' ������� ����������� ��� ����������� �����
			num_EndFolder_0=num_EndFolder_0+1
		End if
	Next

	WriteLog "========== ����� � ����������� ==========" & vbCrLf
	WriteLog "���� ������ " & num_EndFolder & " �����, ��� ����������� � ��� ������." 
	WriteLog "����� ����� ��������� �� �������� - " & num_EndFolder_0 & " ��. (��.���� ����)"
	WriteLog "�������� ���� ���������� - " & num_files & " ������."
	WriteLog "�� ��� ����������� � ������� - " & num_files_copy & " ��., �� ����������� � ���������� ������ - " & err_files_copy & " ��."
	WriteLog "�� ��� ����������� ����� ������ - " & num_files_new & " ��., �� ����������� � ���������� ������ - " & err_files_new & " ��."
	WriteLog "�������� ���� ���������� - " & num_SubFolder & " ��������."
	WriteLog "�� ��� ����������� ����� �������� - " & num_SubFolder_copy & " ��., �� ����������� � ���������� ������ - " & 	err_SubFolder & " ��."
	WriteLog "========== ���������� ������� ��������� ==========" & vbCrLf
	'WScript.Echo "���������� ������� ���������."

Sub CopyFolder(sCopyFolder,sEndCopyFolder)
	' �������� ������� Folder
    Set oFolder = oFSO.GetFolder(sCopyFolder)
	Set oEndCopyFolder = oFSO.GetFolder(sEndCopyFolder)
	' ��������� ��������� ������
    Set colFiles = oFolder.Files
    ' ��������� ������� ����� �� ���������
    For each oFile in colFiles
		WriteLog "���� �������� ����������� �����:" & oFile.Name & vbTab & oFile.DateCreated
		' ������� ����� ����������� ������
		num_files=num_files+1
		' ��������� ���������� ��� ����� ���� � �����, ���� ��� ���, �� ��������. 
		' ���� ����, �� ��������� ��� ������������ � �������� ����� �����, ���� �� �������.
		If oFSO.FileExists(oFSO.BuildPath(oEndCopyFolder, oFile.Name)) Then 
			' ���������� ��������� � ���
			WriteLog "����� ���� ��� ���������� � ����� " & oEndCopyFolder
			' ��������� ��������� ��� ������ ����� �����, ��� ����� ���������� ���� �������� ���� ������
			WriteLog "��������� ������������ �����:"
			' ��������� ������ ���� � ������������ �����
			sFileEnd = oFSO.BuildPath(oEndCopyFolder, oFile.Name)
			' ������� ������ File, ��� ������ � ���� ������
			Set oFileEnd = oFSO.GetFile(sFileEnd)
			' ���������� ���� ��������� ������ 
			If oFileEnd.DateLastModified < oFile.DateLastModified Then
				' ����������� ���� �������� ����������, ������� �������� ��� ����� �����
				WriteLog "����� ����� ��������, �������� ��� �����.     **********" & vbCrLf
				oFSO.CopyFile oFile, sEndCopyFolder & oFile.Name, True
				' �������� �� ������� ������
				if err.Number <> 0 then
					' ������ ��������� �� ������ � ���
					WriteLog "-----> Error # " & CStr(Err.Number) & " " & Err.Description
					' ������� ������
					Err.Clear
					' ������� ������ ��� ������ ������
					err_files_copy=err_files_copy+1
				else
					' ������� ������, ������� ���� �������� �� �����
					num_files_copy=num_files_copy+1				
				End if
			else
				' � ���� ������ ����� ������ ��������, ������ ���������� ������ ������� �����
				WriteLog "����� ���������. ���������� ������." & vbCrLf
			End if
		else
			' ���������� ��������� � ���
			WriteLog "���� ���� ����������� � ����� " & oEndCopyFolder & " ������� ��� ���������." & vbCrLf
			oFSO.CopyFile oFile, sEndCopyFolder & oFile.Name, True
			' �������� �� ������� ������
			if err.Number <> 0 then
				' ������ ��������� �� ������ � ���
				WriteLog "-----> Error # " & CStr(Err.Number) & " " & Err.Description
				' ������� ������
				Err.Clear
				' ������� ������ ��� ����������� ����� ������
				err_files_new=err_files_copy+1
			else
				' ������� ����� ������������� ������
				num_files_new=num_files_copy+1				
			End if
		End if	
	Next
	' ��������� ��� ����� � ��������
	WriteLog "������������ � �������� ��� �������� �� ����� " & oEndCopyFolder & vbCrLf
	' ��������� ��������� ��������
    Set colSubFolders = oFolder.SubFolders
    ' ��������� ������ ��������
    For Each oSubFolder In colSubFolders
		WriteLog "��������� �������� " & oSubFolder
		' ������� ������������ ����� � ��������
		num_SubFolder=num_SubFolder+1	
		' ��������� ���������� ��� ����� �������� � �����, ���� �� ���, �� ��������. 
		' ���� ����, �� ��������� � �������� ������ � ��������.
		If oFSO.FolderExists(oFSO.BuildPath(oEndCopyFolder, oFSO.GetBaseName(oSubFolder.Path))) Then
			' ���������� ��������� � ���
			WriteLog "����� �������� ��� ���������� � ����� " & oEndCopyFolder
			WriteLog "��������� ��� ����� � ���� ��������: "
			' ��������� ������ ���� � ������������ ��������
			sSubFolderEnd = oFSO.BuildPath(oEndCopyFolder, oFSO.GetBaseName(oSubFolder.Path)) & "\" 
			' ���������� ����������� ����� ��������� ����������� ������ - ��������� �������� ���� ����
			CopyFolder oSubFolder, sSubFolderEnd			
			' oLogFile.Writeline
		else
			' ���������� ��������� � ���			
			WriteLog "��� �������� ����������� � ����� " & oEndCopyFolder & " ������� �� ���������." & vbCrLf
			oFSO.CopyFolder oSubFolder, sEndCopyFolder, True
			' �������� �� ������� ������
			if err.Number <> 0 then
				' ������ ��������� �� ������ � ���
				WriteLog "-----> Error # " & CStr(Err.Number) & " " & Err.Description
				' ������� ������
				Err.Clear
				' ������� ������ ��� ����������� ����� ����� � ��������
				err_SubFolder=err_SubFolder+1
			else
				' ������� ����� ������������� ����� � ��������
				num_SubFolder_copy=num_SubFolder_copy+1			
			End if
		End if			
	Next
End Sub

Else
	WriteLog "����� " & StartFolder & " �� ����������"
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