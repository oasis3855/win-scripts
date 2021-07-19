Option Explicit

' **********************************
' テキストエディタを選択して開く.vbs
'
' このファイルは Shift JIS で保存すること
' 別途必要なプログラム
'   DlgDropdownList （スクリプトから呼び出すダイアログボックス単体表示）
'   https://github.com/oasis3855/WindowsScripts/tree/main/DlgDropdownList_VisualC
' **********************************

Dim strArg
Dim strTargetPath
Dim fso
Dim objWshShell

' 引数が無い場合
if Wscript.Arguments.Count < 1 then
    MsgBox("エディタで開きたいファイルを引数に指定してください")
    WScript.Quit
End If

' プログラムの引数（Explorerの送る(Send To)により、フルパスのファイル名が格納されている
strArg = Wscript.Arguments(0)

' 引数1で与えられたものが、ファイルかどうかを判定し、ファイル以外の場合は終了する
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FolderExists(strArg) Then
    MsgBox("引数で指定されたのはディレクトリです。処理対象外のため終了します")
    WScript.Quit
ElseIf fso.FileExists(strArg) <> True Then
    MsgBox("引数で指定されたのはファイルではありません。処理対象外のため終了します")
    WScript.Quit
End If


Set objWshShell = CreateObject("WScript.Shell")

' エディタの種類を選択するダイアログボックスを表示し、選択した値をExitCodeとして得る
' Runメソッドの第2引数は実行するウインドウの状態、第3引数はプロセスの修了を待つかどうかの設定
Dim numExitCode
numExitCode = objWshShell.Run("DlgDropdownList.exe" & " " & "起動するテキストエディタを選択してください Notepad Terapad サクラエディタ ""Visual Studio Code"" Bzバイナリエディタ", 1, True)

' 選択したエディタで、指定されたファイル（strArg）を開く
' Execメソッドは、プロセスの終了を待たずに制御を返す
If numExitCode = 1 Then
    objWshShell.Exec("notepad.exe" & " " & strArg)
ElseIf numExitCode = 2 Then
    objWshShell.Exec("C:\Program Files\OnlineSoftware\TextEditor\TeraPad\TeraPad.exe" & " " & strArg)
ElseIf numExitCode = 3 Then
    objWshShell.Exec("C:\Program Files\OnlineSoftware\TextEditor\sakura\sakura.exe" & " " & strArg)
ElseIf numExitCode = 4 Then
    objWshShell.Exec("C:\Program Files\Microsoft VS Code\Code.exe" & " " & strArg)
ElseIf numExitCode = 5 Then
    objWshShell.Exec("C:\Program Files\OnlineSoftware\TextEditor\Bz\Bz.exe" & " " & strArg)
End If

