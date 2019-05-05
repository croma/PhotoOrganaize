' 綺麗じゃないけど動けばええやんな？（暴論）
dim fso, crDir, WshShell, dirObj, fo, fldrPath, dlm

' カレントディレクトリ取得
Set WshShell = WScript.CreateObject("WScript.Shell")
crDir = WshShell.CurrentDirectory
set WshShell = nothing

Set fso = CreateObject("Scripting.FileSystemObject")
set dirObj = fso.GetFolder(crDir)
for each fo in dirObj.files
	if fo.Name = WScript.ScriptName then
		' 自分ならなにもしない
	else
		' 最終更新日の取得
		dlm = fo.DateLastModified

		' 移動先のフォルダはなければ作る。最終更新日をMMyyで。
		fldrPath= crDir & "\" & Right("0" & month(dlm) , 2) & Right("0" & Day(dlm) , 2) & "\"
		If not (fso.FolderExists(fldrPath)) Then
			call fso.CreateFolder(fldrPath)
		End If

		' 移動
		call fso.movefile(fo.Name, fldrPath)
	end if
next

set fso = nothing
set fo = nothing
