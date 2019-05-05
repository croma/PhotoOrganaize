' ImageFileあるのでWinVista以降でないと動かないはず。動作確認はWin7
' 綺麗じゃないけど動けばええやんな？（暴論）
dim fso, crDir, WshShell, dirObj, fo, fldrPath, dlm, rawExp, rawName, imgFile, takeDate

' === 使うときに書き換えてください ここから ===
rawExp = "CR2" ' RAW画像の拡張子。大文字小文字区別するっぽいので注意
' === 使うときに書き換えてください ここまで ===

' カレントディレクトリ取得
Set WshShell = WScript.CreateObject("WScript.Shell")
crDir = WshShell.CurrentDirectory
set WshShell = nothing

Set fso = CreateObject("Scripting.FileSystemObject")
Set imgFile = CreateObject("Wia.ImageFile")
set dirObj = fso.GetFolder(crDir)
for each fo in dirObj.files
	if Right(lcase(fo.name), 3) = "jpg" then
		imgFile.LoadFile crDir & "\" & fo.name
		' Exifの撮影日（なければ最終更新日）の取得
                takeDate = imgFile.Properties("36867")
		if takeDate <> "" then
			' YYYY:MM:DD HH:MM:SS
			dlm = Mid(takeDate, 6, 2) & Mid(takeDate, 9, 2)
                else
			' 最終更新日は日付型
			dlm = Right("0" & month(fo.DateLastModified) , 2) & Right("0" & Day(fo.DateLastModified) , 2)
                end if
		' Wia.ImageFile.LoadFileに対応するCloseがなさそう、リーク不安。

		' 移動先のフォルダはなければ作る。最終更新日をMMyyで。
		fldrPath= crDir & "\" & dlm & "\"
		if not (fso.FolderExists(fldrPath)) Then
			call fso.CreateFolder(fldrPath)
		end if

		' 移動
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
