' ImageFile����̂�WinVista�ȍ~�łȂ��Ɠ����Ȃ��͂��B����m�F��Win7
' �Y�킶��Ȃ����Ǔ����΂������ȁH�i�\�_�j
dim fso, crDir, WshShell, dirObj, fo, fldrPath, dlm, rawExp, rawName, imgFile, takeDate

' === �g���Ƃ��ɏ��������Ă������� �������� ===
rawExp = "CR2" ' RAW�摜�̊g���q�B�啶����������ʂ�����ۂ��̂Œ���
' === �g���Ƃ��ɏ��������Ă������� �����܂� ===

' �J�����g�f�B���N�g���擾
Set WshShell = WScript.CreateObject("WScript.Shell")
crDir = WshShell.CurrentDirectory
set WshShell = nothing

Set fso = CreateObject("Scripting.FileSystemObject")
Set imgFile = CreateObject("Wia.ImageFile")
set dirObj = fso.GetFolder(crDir)
for each fo in dirObj.files
	if Right(lcase(fo.name), 3) = "jpg" then
		imgFile.LoadFile crDir & "\" & fo.name
		' Exif�̎B�e���i�Ȃ���΍ŏI�X�V���j�̎擾
                takeDate = imgFile.Properties("36867")
		if takeDate <> "" then
			' YYYY:MM:DD HH:MM:SS
			dlm = Mid(takeDate, 6, 2) & Mid(takeDate, 9, 2)
                else
			' �ŏI�X�V���͓��t�^
			dlm = Right("0" & month(fo.DateLastModified) , 2) & Right("0" & Day(fo.DateLastModified) , 2)
                end if
		' Wia.ImageFile.LoadFile�ɑΉ�����Close���Ȃ������A���[�N�s���B

		' �ړ���̃t�H���_�͂Ȃ���΍��B�ŏI�X�V����MMyy�ŁB
		fldrPath= crDir & "\" & dlm & "\"
		if not (fso.FolderExists(fldrPath)) Then
			call fso.CreateFolder(fldrPath)
		end if

		' �ړ�
		rawName = left(fo.Name, len(fo.Name) - 3) & rawExp
		call fso.movefile(fo.Name, fldrPath)
		if fso.FileExists(crDir & "\" & rawName) then
			call fso.movefile(rawName, fldrPath)
		end if
	end if
next

set imgFile = nothing
set fso = nothing
set fo = nothing
