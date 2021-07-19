Option Explicit

' **********************************
' ファイル名・更新日時変更.vbs
'
' このファイルは Shift JIS で保存すること
' 別途必要なプログラム
'   DlgDropdownList （スクリプトから呼び出すダイアログボックス単体表示）
'   https://github.com/oasis3855/WindowsScripts/tree/main/DlgDropdownList_VisualC
' **********************************

Call Main()

Sub Main()

    Dim modeFolder      ' ディレクトリ内全ファイル対象とするモードの場合は True
    Dim strTargetPath   ' 対象ファイル または ディレクトリ

    ' このスクリプトの引数は1個かどうかのチェック
    if Wscript.Arguments.Count <> 1 then
        MsgBox("引数（対象とするファイル）を1つ指定してください")
        WScript.Quit
    End If

    ' 引数で指定されたものが、ファイルか、ディレクトリかチェックする
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim strArg
    strArg = Wscript.Arguments(0)
    If fso.FolderExists(strArg) Then
        ' 引数はディレクトリ
        modeFolder = True
        strTargetPath = strArg
    ElseIf fso.FileExists(strArg) Then
        ' 引数はファイル
        modeFolder = False
    Else
        MsgBox(strArg & " は存在しません")
        WScript.Quit
    End If

    ' ファイルが引数で指定されているとき、ファイル単体か、ファイルのあるディレクトリ全体かの選択をする
    Dim ans
    If modeFolder = False Then
        ans = MsgBox("処理対象の選択" & vbCRLF & vbCRLF & _
                    "[YES] ファイル単体を対象 = " & strArg & vbCRLF & _
                    "[NO] ディレクトリ内全体を対象 = " & fso.GetParentFolderName(strArg) & vbCRLF & vbCRLF & _
                    "YES or No を選択してください" & vbCRLF & vbCRLF , vbYesNoCancel)
        If ans=vbYes Then
            strTargetPath = strArg
        ElseIf ans = vbNo Then
            modeFolder = True
            strTargetPath = fso.GetParentFolderName(strArg)
        Else
            MsgBox("キャンセルしました")
            WScript.Quit
        End If
    End If
    set fso = Nothing

    ' リストボックスのダイアログで、機能の選択を行う
    Dim numUserSelect
    numUserSelect = SelectTaskDialog(strTargetPath, modeFolder)

    ' ファイル名の変更の場合、リストボックスのダイアログで、対象がベースネームか拡張子かを選択する
    Dim modeBasenameExt
    Dim retNum
    If numUserSelect = 1 Or numUserSelect = 2 Then
        retNum = SelectBaseExtDialog(strTargetPath, "small")
        If retNum = 1 Then
            modeBasenameExt = "basename"
        ElseIf retNum = 2 Then
            modeBasenameExt = "ext"
        ElseIf retNum = 3 Then
            modeBasenameExt = "basename.ext"
        Else
            MsgBox("ユーザによってキャンセルされました")
            Exit Sub
        End If
    End If

    Dim counter     ' 処理完了したファイルの個数（結果表示用）
    counter = 0
    If numUserSelect = 1 then
        counter = ChangeFnames(strTargetPath, modeFolder, "small", modeBasenameExt)
        MsgBox(counter & " 個のファイル名を変更しました")
    ElseIf numUserSelect = 2 then
        counter = ChangeFnames(strTargetPath, modeFolder, "LARGE", modeBasenameExt)
        MsgBox(counter & " 個のファイル名を変更しました")
    ElseIf numUserSelect = 3 then
        counter = TouchFiles(strTargetPath, modeFolder, Now)
        MsgBox(counter & " 個のファイルの更新日時を現在日時に変更しました")
    ElseIf numUserSelect = 4 then
        Dim SelectedDateTime
        SelectedDateTime = InputBox("更新日時のユーザ入力", "対象ファイル" & strTargetPath, Now)
        counter = TouchFiles(strTargetPath, modeFolder, SelectedDateTime)
        MsgBox(counter & " 個のファイルの更新日時を " & SelectedDateTime & " に変更しました")
    Else
        MsgBox("ユーザによってキャンセルされました")
    End If

End Sub

' 機能選択ダイアログ（リストボックスより選択）を表示する
Function SelectTaskDialog(strTargetPath, modeFolder)
    SelectTaskDialog = 0    ' デフォルトの戻り値

    Dim strMsg
    If modeFolder = True Then strMsg = "[ディレクトリ]  " Else strMsg = "[ファイル単体]  "

    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    SelectTaskDialog = objWshShell.Run("DlgDropdownList.exe" & " " & """" & strMsg & strTargetPath & "  に対して行う処理を選択してください"" " & _
                        "ファイル名を小文字に変更 ファイル名を大文字に変更 最終更新日を現在日時に変更 最終更新日を指定した日時に変更", 1, True)
End Function

' リネーム対象選択ダイアログ（リストボックスより選択）を表示する
Function SelectBaseExtDialog(strTargetPath, modeLargeSmall)
    SelectBaseExtDialog = 0    ' デフォルトの戻り値

    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    Dim strLargeSmall
    If modeLargeSmall = "small" Then
        strLargeSmall = "小文字化"
    Else
        strLargeSmall = "大文字化"
    End If
    Dim retNum
    SelectBaseExtDialog = objWshShell.Run("DlgDropdownList.exe" & " " & """" & strTargetPath & " のファイル名の" & strLargeSmall & "の対象を選択してください"" " & _
                        "ベースネームのみ対象 拡張子のみ対象 ベースネームと拡張子の両方を対象", 1, True)
End Function

' 指定したファイル（またはディレクトリ内の全ファイル）の更新日時を変更する
Function TouchFiles(strTargetPath, modeFolder, SelectedDateTime)
    ' 戻り値の初期設定（処理が成功したファイル数）
    TouchFiles = 0

    ' 単体ファイルのとき
    If modeFolder = False Then
        If TouchFile(strTargetPath, SelectedDateTime) Then
            ' ファイル1個の処理が成功した時の戻り値
            TouchFiles = 1
        End If
        Exit Function
    End If

    ' 日時文字列のフォーマットチェック
    If IsDate(SelectedDateTime) = False Then
        TouchFiles = 0
        MsgBox("日時の書式が規定外です : " & SelectedDateTime)
        exit Function
    End If

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim objShell
    Set objShell = CreateObject("Shell.Application")

    Dim objFiles
    Dim objFile
    ' 指定されたディレクトリ内の全ファイル一覧
    Set objFiles = fso.GetFolder(strTargetPath).files
    ' ファイル1個ずつ処理
    For Each objFile In objFiles
        If TouchFile(objFile, SelectedDateTime) Then
            ' ファイル1この処理が成功した時、戻り値（成功ファイル数カウンタ）をインクリメントする
            TouchFiles = TouchFiles + 1
        End If
    Next

    set objFile = Nothing
    set objFiles = Nothing

End Function

' 指定したファイル単体の更新日時を変更する
Function TouchFile(strTargetPath, SelectedDateTime)
    ' 戻り値の初期設定（成功時にTrueを返す）
    TouchFile = True

    ' 日時文字列のフォーマットチェック
    If IsDate(SelectedDateTime) = False Then
        TouchFile = False
        MsgBox("日時の書式が規定外です : " & SelectedDateTime)
        exit Function
    End If

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim objShell
    Set objShell = CreateObject("Shell.Application")
    Dim objFolder
    Set objFolder = objShell.NameSpace(fso.GetParentFolderName(strTargetPath))

'   MsgBox("strTargetPath = " & strTargetPath & vbCRLF & _
'           "fso.GetParentFolderName(strTargetPath) = " & fso.GetParentFolderName(strTargetPath) & vbCRLF & _
'           "fso.GetFileName(strTargetPath) = " & fso.GetFileName(strTargetPath) & vbCRLF & _
'           "SelectedDateTime = " & SelectedDateTime)

    ' FolderItem が正しく得られた場合に、更新日時を変更する
    If (Not objFolder.Items.Item(fso.GetFileName(strTargetPath)) is Nothing) then
        objFolder.Items.Item(fso.GetFileName(strTargetPath)).ModifyDate = SelectedDateTime
    Else
        TouchFile = False
    End If

    set objFolder = Nothing

End Function

' 指定したファイル（またはディレクトリ内の全ファイル）のファイル名を、大文字化・小文字化する
Function ChangeFnames(strTargetPath, modeFolder, modeLargeSmall, modeBasenameExt)
    ' 戻り値の初期設定（処理が成功したファイル数）
    ChangeFnames = 0
    ' 単体ファイルの時
    If modeFolder = False Then
        If ChangeFname(strTargetPath, modeLargeSmall, modeBasenameExt) = True then
            ChangeFnames = ChangeFnames + 1
        End If
        Exit Function
    End If

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim objShell
    Set objShell = CreateObject("Shell.Application")

    Dim objFiles
    Dim objFile
    ' 指定されたディレクトリ内の全ファイル一覧
    Set objFiles = fso.GetFolder(strTargetPath).files
    ' ファイル1個ずつ処理
    For Each objFile In objFiles
        If ChangeFname(objFile, modeLargeSmall, modeBasenameExt) = True then
            ChangeFnames = ChangeFnames + 1
        End If
    Next

    set objFile = Nothing
    set objFiles = Nothing



End Function

' 指定したファイル単体のファイル名を、大文字化・小文字化する
Function ChangeFname(strTargetPath, modeLargeSmall, modeBasenameExt)
    ' 戻り値の初期設定（成功時にTrueを返す）
    ChangeFname = False

    ' Scripting.FileSystemObject の File.Name でのファイル名変更は、大文字小文字区別ないためエラーとなる
    ' そのため、Shell.Application の FolderItem.Name でファイル名変更する

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim objShell
    Set objShell = CreateObject("Shell.Application")

    ' 対象ファイルが存在することを確認
    If fso.FileExists(strTargetPath) <> True Then
        ChangeFname = False
        MsgBox("対象ファイルが見つからない : " & strTargetPath)
        Exit Function
    End If

    ' パス名を分解し、ディレクトリ、ファイル本体、拡張子の文字列に格納
    Dim dirname
    Dim filename
    Dim basename
    Dim ext
    dirname = fso.GetParentFolderName(strTargetPath)
    filename = fso.GetFileName(strTargetPath)
    basename = fso.GetBaseName(strTargetPath)
    ext = fso.GetExtensionName(strTargetPath)

    Dim objFolder
    Set objFolder = objShell.NameSpace(dirname)
    Dim objFolderitem
    Set objFolderitem = objFolder.ParseName(filename)
    ' 変更後のファイル名の作成
    Dim filenameNew
    If modeLargeSmall = "small" and InStr(modeBasenameExt, "basename") > 0 Then
        basename = LCase(basename)
    End If
    If modeLargeSmall = "small" and InStr(modeBasenameExt, "ext") > 0 Then
        ext = LCase(ext)
    End If
    If modeLargeSmall = "LARGE" and InStr(modeBasenameExt, "basename") > 0 Then
        basename = UCase(basename)
    End If
    If modeLargeSmall = "LARGE" and InStr(modeBasenameExt, "ext") > 0 Then
        ext = UCase(ext)
    End If
    filenameNew = basename & "." & ext

    ' ファイル名変更後と、同一名のファイル（大文字小文字判別）、同一名フォルダが存在しないことを確認
    If IsFileExistCaseSensitive(dirname & "\" & filenameNew) = True Or fso.FolderExists(dirname & "\" & filenameNew) = True Then
        ChangeFname = False
        ' デバッグ時のメッセージ
        ' MsgBox("同一名のファイルが存在 : " & filenameNew)
        Exit Function
    End If
    ' ファイル名の変更
    objFolderitem.Name = filenameNew
    ' 処理成功時の戻り値
    ChangeFname = True
End Function

' ファイル名の大文字小文字を区別して、ファイルが存在するかチェックする
Function IsFileExistCaseSensitive(strTargetPath)
    ' 戻り値の初期設定
    IsFileExistCaseSensitive = False

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim objShell
    Set objShell = CreateObject("Shell.Application")
    Dim objFolder
    Set objFolder = objShell.NameSpace(fso.GetParentFolderName(strTargetPath))
    Dim objFolderitem
    Set objFolderitem = objFolder.ParseName(fso.GetFileName(strTargetPath))

    If objFolderitem.Name = fso.GetFileName(strTargetPath) Then
        IsFileExistCaseSensitive = True
    End If
End Function
