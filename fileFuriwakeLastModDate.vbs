' �Y�킶��Ȃ����Ǔ����΂������ȁH�i�\�_�j
dim fso, crDir, WshShell, dirObj, fo, fldrPath, dlm

' �J�����g�f�B���N�g���擾
Set WshShell = WScript.CreateObject("WScript.Shell")
crDir = WshShell.CurrentDirectory
set WshShell = nothing

Set fso = CreateObject("Scripting.FileSystemObject")
set dirObj = fso.GetFolder(crDir)
for each fo in dirObj.files
	if fo.Name = WScript.ScriptName then
		' �����Ȃ�Ȃɂ����Ȃ�
	else
		' �ŏI�X�V���̎擾
		dlm = fo.DateLastModified

		' �ړ���̃t�H���_�͂Ȃ���΍��B�ŏI�X�V����MMyy�ŁB
		fldrPath= crDir & "\" & Right("0" & month(dlm) , 2) & Right("0" & Day(dlm) , 2) & "\"
		If not (fso.FolderExists(fldrPath)) Then
			call fso.CreateFolder(fldrPath)
		End If

		' �ړ�
		call fso.movefile(fo.Name, fldrPath)
	end if
next

set fso = nothing
set fo = nothing
