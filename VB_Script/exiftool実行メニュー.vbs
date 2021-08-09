Option Explicit

' **********************************
' exiftool実行メニュー.vbs
'
' このファイルは Shift JIS で保存すること
' 別途必要なプログラム
'   ExifTool by Phil Harvey（exiftool.exeをPATHの通ったディレクトリに置いてください）
'   https://exiftool.org/
'   DlgDropdownList （スクリプトから呼び出すダイアログボックス単体表示）
'   https://github.com/oasis3855/WindowsScripts/tree/main/DlgDropdownList_VisualC
' **********************************

Call Main()

Sub Main()

    Dim strPath
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 引数が無い場合
    if Wscript.Arguments.Count < 1 then
        MsgBox("開きたいファイルを引数に指定してください")
        WScript.Quit
    End If

    ' プログラムの引数（Explorerの送る(Send To)により、フルパスのファイル名が格納されている
    strPath = Wscript.Arguments(0)

    ' 引数がファイルまたはディレクトリでない場合
    If fso.FolderExists(strPath) = False and fso.FileExists(strPath) = False Then
        MsgBox(strPath & " は存在しません")
        WScript.Quit
    End If

    Dim numUserSelect
    numUserSelect = SelectTaskDialog(strPath)

    If numUserSelect = 1 Then
        ' Exifタグ情報を出力
        ' 引数1で与えられたものが、ファイルかどうかを判定し、ファイル以外の場合は終了する
        If CheckFileOnly(strPath) <> True Then
            MsgBox("引数で指定されたのはディレクトリなどで、ファイルではありません。処理対象外のため終了します")
            WScript.Quit
        End If
        ShowExif(strPath)
    ElseIf numUserSelect = 2 Then
        ' ファイルのタイムスタンプをExif日時に変更
        ' ファイル単体か、ディレクトリ全体かの選択
        strPath = SelectFileOrDir(strPath)
        RestoreChangeTimeStamp(strPath)
    ElseIf numUserSelect = 3 Then
        ' Gimp編集前にMakeタグのPENTAXをPentax Corp.に書き換え
        ' ファイル単体か、ディレクトリ全体かの選択
        strPath = SelectFileOrDir(strPath)
        ChangePentaxMakeTag(strPath)
    ElseIf numUserSelect = 4 Then
        ' Gimp編集後にSoftwareとModifyDateの書き戻し修正
        ' ファイル単体か、ディレクトリ全体かの選択
        strPath = SelectFileOrDir(strPath)
        RestoreGimpChangedTag(strPath)
    ElseIf numUserSelect = 5 Then
        ' ファイル名をExif:CreateDateに変更
        ' ファイル単体か、ディレクトリ全体かの選択
        strPath = SelectFileOrDir(strPath)
        RenameFilename(strPath)
    Else
        MsgBox("キャンセルしました")
    End If
    Set fso = Nothing

End Sub

' 機能選択ダイアログ
Function SelectTaskDialog(strPath)
    SelectTaskDialog = 0    ' デフォルトの戻り値

    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    SelectTaskDialog = objWshShell.Run("DlgDropdownList.exe" & " " & """" & strPath & " に対して行う処理を選択してください""" & _
                        " Exifタグ情報表示" & _
                        " ファイルのタイムスタンプをExif日時に変更" & _
                        " ""Gimp編集前にMakeタグのPENTAXをPentax Corp.に書き換え""" & _
                        " Gimp編集後にSoftwareとModifyDateの書き戻し修正" & _
                        " ファイル名をExif:CreateDateに変更", 1, True)

    Set objWshShell = Nothing
End Function

' 「Gimp編集後にSoftwareとModifyDateの書き戻し修正」での実行対象カメラメーカー選択
Function SelectCameraMaker(strPath)
    SelectCameraMaker = 0    ' デフォルトの戻り値

    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    SelectCameraMaker = objWshShell.Run("DlgDropdownList.exe" & " " & "修正するMakeタグを選択してください" & _
                        " ""PENTAX K-x""" & _
                        " ""SONY DSC-RX100""" & _
                        " ""SONY DSC-WX50""" & _
                        " ""Huawei P10 Lite (WAS-LX2J)""" & _
                        " ""ASUS ZenFone Max M2 (msm8953_64-user)""" & _
                        " ""Xiaomi Redmi 9T (M2010J19SG)""", 1, True)

    Set objWshShell = Nothing
End Function

' exif情報を表示する
Sub ShowExif(strJpegFilepath)
    ' 一時ファイルのフルパス名を生成する
    Dim TempFilePath
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    TempFilePath = fso.GetSpecialFolder(2) & "\" & fso.GetTempName
    Set fso = Nothing

    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")

    ' exiftoolを実行し、exifデータを読み出す
    Dim objExec
    Set objExec = objWshShell.Exec("exiftool -G " & strJpegFilepath)
    'Set objExec = objWshShell.Exec("%comspec% /c exiftool -G " & strArg & " > " & TempFilePath)

    ' 一時ファイルに、先ほど読みだされたexifデータを書き込む
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tso
    Set tso = fso.OpenTextFile(TempFilePath, 2, true)
    tso.Write(objExec.StdOut.ReadAll)
    tso.Close

    Set tso = Nothing
    Set objExec = Nothing

    ' 一時ファイル（exifデータが書き込まれている）をメモ帳で開く
'    objWshShell.Run("notepad.exe" & " " & TempFilePath, 1, 1)
    objWshShell.Run "notepad.exe" & " " & TempFilePath, 1, True

    ' 一時ファイルを削除する
    fso.DeleteFile(TempFilePath)
    Set fso = Nothing
End Sub

' ファイルのタイムスタンプをExif:DateTimeOriginalに変更する
Sub RestoreChangeTimeStamp(strPath)
    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    Dim strCommand : strCommand = "exiftool ""-FileModifyDate<DateTimeOriginal"" "
    ' exiftoolを実行
    Dim objExec
    Set objExec = objWshShell.Exec(strCommand & strPath)
    MsgBox("実行コマンド : " & strCommand & strPath & vbCRLF & vbCRLF & objExec.StdOut.ReadAll)
End Sub

' Gimp編集前にMakeタグのPENTAXをPentax Corp.に書き換え
Sub ChangePentaxMakeTag(strPath)
    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    Dim strCommand : strCommand = "exiftool -if ""$Exif:Make eq 'PENTAX'"" -overwrite_original -preserve -Exif:Make=""Pentax Corp."" "
    ' exiftoolを実行
    Dim objExec
    Set objExec = objWshShell.Exec(strCommand & strPath)
    MsgBox("実行コマンド : " & strCommand & strPath & vbCRLF & vbCRLF & objExec.StdOut.ReadAll)

    Set objExec = Nothing
    Set objWshShell = Nothing
End Sub

' Gimp編集後に破壊されたSoftwareとModifyDateの書き戻し修正
Sub RestoreGimpChangedTag(strPath)
    Dim numUserSelect
    numUserSelect = SelectCameraMaker(strPath)

    Dim strCommand
    If numUserSelect = 1 Then
        ' PENTAX K-x
        strCommand = "exiftool -if ""$Exif:Model =~ /PENTAX K-x/i and $Exif:Software =~ /Gimp/i"" -overwrite_original -preserve -Exif:Software=""K-x Ver 1.03"" ""-Exif:ModifyDate<$Exif:CreateDate"" "
    ElseIf  numUserSelect = 2 Then
        ' DSC-RX100
        strCommand = "exiftool -if ""$Exif:Model =~ /DSC-RX100/i and $Exif:Software =~ /Gimp/i"" -overwrite_original -preserve -Exif:Software=""DSC-RX100 v1.10"" ""-Exif:ModifyDate<$Exif:CreateDate"" "
    ElseIf  numUserSelect = 3 Then
        ' DSC-WX50
        strCommand = "exiftool -if ""$Exif:Model =~ /DSC-WX50/i and $Exif:Software =~ /Gimp/i"" -overwrite_original -preserve -Exif:Software=""DSC-WX50 v1.00"" ""-Exif:ModifyDate<$Exif:CreateDate"" "
    ElseIf  numUserSelect = 4 Then
        ' Huawei P10 Lite WAS-LX2J
        strCommand = "exiftool -if ""$Exif:Model =~ /WAS-LX2J/i and $Exif:Software =~ /Gimp/i"" -overwrite_original -preserve -Exif:Software=""WAS-LX2JC635B106"" ""-Exif:ModifyDate<$Exif:CreateDate"" "
    ElseIf  numUserSelect = 5 Then
        ' ASUS ZenFone Max M2
        strCommand = "exiftool -if ""$Exif:Model =~ /ASUS_X01AD/i and $Exif:Software =~ /Gimp/i"" -overwrite_original -preserve -Exif:Software=""msm8953_64-user 9 WW_Phone-202011271133 78"" ""-Exif:ModifyDate<$Exif:CreateDate"" "
    ElseIf  numUserSelect = 6 Then
        ' Xiaomi Redmi 9T
        strCommand = "exiftool -if ""$Exif:Model =~ /M2010J19SG/i and $Exif:Software =~ /Gimp/i"" -overwrite_original -preserve -Exif:Software=""Xiaomi Redmi 9T"" ""-Exif:ModifyDate<$Exif:CreateDate"" "
    Else
        MsgBox("キャンセルしました")
        Exit Sub
    End If

    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    ' exiftoolを実行
    Dim objExec
    Set objExec = objWshShell.Exec(strCommand & strPath)
    MsgBox("実行コマンド : " & strCommand & strPath & vbCRLF & vbCRLF & objExec.StdOut.ReadAll)

    Set objExec = Nothing
    Set objWshShell = Nothing
End Sub

' ファイル名をExif:CreateDateに変更
Sub RenameFilename(strPath)

    Dim strCommand
    Dim ans
    ans = MsgBox("実行/テストの選択" & vbCRLF & vbCRLF & _
                "[YES] ファイル名を変更 " & vbCRLF & _
                "[NO] シミュレーションのみ " & vbCRLF & vbCRLF & _
                "YES or No を選択してください" & vbCRLF & vbCRLF , vbYesNoCancel)
    If ans=vbYes Then
        strCommand = "exiftool ""-FileName<$Exif:CreateDate"" -d %Y-%m-%d-%H-%M-%S_%%f.%%e "
    ElseIf ans = vbNo Then
        strCommand = "exiftool ""-TestName<$Exif:CreateDate"" -d %Y-%m-%d-%H-%M-%S_%%f.%%e "
    Else
        MsgBox("キャンセルしました")
        WScript.Quit
    End If

    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    ' exiftoolを実行
    Dim objExec
    Set objExec = objWshShell.Exec(strCommand & strPath)
    MsgBox("実行コマンド : " & strCommand & strPath & vbCRLF & vbCRLF & objExec.StdOut.ReadAll)

    Set objExec = Nothing
    Set objWshShell = Nothing
End Sub

' strPathがファイル名かどうかを判定する（ファイル名：True，ディレクトリ名：False）
Function CheckFileOnly(strPath)
    Dim fso
    CheckFileOnly = False  ' 戻り値

    ' 引数で与えられたフルパス名が、ファイルかどうかを判定し、ファイルの場合はTrueを返す
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(strPath) Then
        ' 引数で指定されたのはディレクトリ
        CheckFileOnly = False  ' 戻り値
        Exit Function
    ElseIf fso.FileExists(strPath) <> True Then
        ' 引数で指定されたのはファイルではない
        CheckFileOnly = False  ' 戻り値
        Exit Function
    End If
    Set fso = Nothing

    CheckFileOnly = True  ' 戻り値
End Function

' ファイル単体か、ファイルのあるディレクトリ全体化を選択し、それに従ったパス名を返す
Function SelectFileOrDir(strPath)
    SelectFileOrDir = strPath ' 戻り値
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 引数がディレクトリの場合は、そのまま終了する（ディレクトリ名を返す）
    If fso.FolderExists(strPath) Then Exit Function

    Dim ans
    ans = MsgBox("処理対象の選択" & vbCRLF & vbCRLF & _
                "[YES] ファイル単体を対象 = " & strPath & vbCRLF & _
                "[NO] ディレクトリ内全体を対象 = " & fso.GetParentFolderName(strPath) & vbCRLF & vbCRLF & _
                "YES or No を選択してください" & vbCRLF & vbCRLF , vbYesNoCancel)
    If ans=vbYes Then
        SelectFileOrDir = strPath ' 戻り値
    ElseIf ans = vbNo Then
        ' 与えられたファイルの親ディレクトリ名を返す
        SelectFileOrDir = fso.GetParentFolderName(strPath) ' 戻り値
    Else
        MsgBox("キャンセルしました")
        WScript.Quit
    End If
    Set fso = Nothing
End Function
