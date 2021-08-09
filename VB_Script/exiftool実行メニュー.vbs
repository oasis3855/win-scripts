Option Explicit

' **********************************
' exiftool���s���j���[.vbs
'
' ���̃t�@�C���� Shift JIS �ŕۑ����邱��
' �ʓr�K�v�ȃv���O����
'   ExifTool by Phil Harvey�iexiftool.exe��PATH�̒ʂ����f�B���N�g���ɒu���Ă��������j
'   https://exiftool.org/
'   DlgDropdownList �i�X�N���v�g����Ăяo���_�C�A���O�{�b�N�X�P�̕\���j
'   https://github.com/oasis3855/WindowsScripts/tree/main/DlgDropdownList_VisualC
' **********************************

Call Main()

Sub Main()

    Dim strPath
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' �����������ꍇ
    if Wscript.Arguments.Count < 1 then
        MsgBox("�J�������t�@�C���������Ɏw�肵�Ă�������")
        WScript.Quit
    End If

    ' �v���O�����̈����iExplorer�̑���(Send To)�ɂ��A�t���p�X�̃t�@�C�������i�[����Ă���
    strPath = Wscript.Arguments(0)

    ' �������t�@�C���܂��̓f�B���N�g���łȂ��ꍇ
    If fso.FolderExists(strPath) = False and fso.FileExists(strPath) = False Then
        MsgBox(strPath & " �͑��݂��܂���")
        WScript.Quit
    End If

    Dim numUserSelect
    numUserSelect = SelectTaskDialog(strPath)

    If numUserSelect = 1 Then
        ' Exif�^�O�����o��
        ' ����1�ŗ^����ꂽ���̂��A�t�@�C�����ǂ����𔻒肵�A�t�@�C���ȊO�̏ꍇ�͏I������
        If CheckFileOnly(strPath) <> True Then
            MsgBox("�����Ŏw�肳�ꂽ�̂̓f�B���N�g���ȂǂŁA�t�@�C���ł͂���܂���B�����ΏۊO�̂��ߏI�����܂�")
            WScript.Quit
        End If
        ShowExif(strPath)
    ElseIf numUserSelect = 2 Then
        ' �t�@�C���̃^�C���X�^���v��Exif�����ɕύX
        ' �t�@�C���P�̂��A�f�B���N�g���S�̂��̑I��
        strPath = SelectFileOrDir(strPath)
        RestoreChangeTimeStamp(strPath)
    ElseIf numUserSelect = 3 Then
        ' Gimp�ҏW�O��Make�^�O��PENTAX��Pentax Corp.�ɏ�������
        ' �t�@�C���P�̂��A�f�B���N�g���S�̂��̑I��
        strPath = SelectFileOrDir(strPath)
        ChangePentaxMakeTag(strPath)
    ElseIf numUserSelect = 4 Then
        ' Gimp�ҏW���Software��ModifyDate�̏����߂��C��
        ' �t�@�C���P�̂��A�f�B���N�g���S�̂��̑I��
        strPath = SelectFileOrDir(strPath)
        RestoreGimpChangedTag(strPath)
    ElseIf numUserSelect = 5 Then
        ' �t�@�C������Exif:CreateDate�ɕύX
        ' �t�@�C���P�̂��A�f�B���N�g���S�̂��̑I��
        strPath = SelectFileOrDir(strPath)
        RenameFilename(strPath)
    Else
        MsgBox("�L�����Z�����܂���")
    End If
    Set fso = Nothing

End Sub

' �@�\�I���_�C�A���O
Function SelectTaskDialog(strPath)
    SelectTaskDialog = 0    ' �f�t�H���g�̖߂�l

    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    SelectTaskDialog = objWshShell.Run("DlgDropdownList.exe" & " " & """" & strPath & " �ɑ΂��čs��������I�����Ă�������""" & _
                        " Exif�^�O���\��" & _
                        " �t�@�C���̃^�C���X�^���v��Exif�����ɕύX" & _
                        " ""Gimp�ҏW�O��Make�^�O��PENTAX��Pentax Corp.�ɏ�������""" & _
                        " Gimp�ҏW���Software��ModifyDate�̏����߂��C��" & _
                        " �t�@�C������Exif:CreateDate�ɕύX", 1, True)

    Set objWshShell = Nothing
End Function

' �uGimp�ҏW���Software��ModifyDate�̏����߂��C���v�ł̎��s�ΏۃJ�������[�J�[�I��
Function SelectCameraMaker(strPath)
    SelectCameraMaker = 0    ' �f�t�H���g�̖߂�l

    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    SelectCameraMaker = objWshShell.Run("DlgDropdownList.exe" & " " & "�C������Make�^�O��I�����Ă�������" & _
                        " ""PENTAX K-x""" & _
                        " ""SONY DSC-RX100""" & _
                        " ""SONY DSC-WX50""" & _
                        " ""Huawei P10 Lite (WAS-LX2J)""" & _
                        " ""ASUS ZenFone Max M2 (msm8953_64-user)""" & _
                        " ""Xiaomi Redmi 9T (M2010J19SG)""", 1, True)

    Set objWshShell = Nothing
End Function

' exif����\������
Sub ShowExif(strJpegFilepath)
    ' �ꎞ�t�@�C���̃t���p�X���𐶐�����
    Dim TempFilePath
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    TempFilePath = fso.GetSpecialFolder(2) & "\" & fso.GetTempName
    Set fso = Nothing

    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")

    ' exiftool�����s���Aexif�f�[�^��ǂݏo��
    Dim objExec
    Set objExec = objWshShell.Exec("exiftool -G " & strJpegFilepath)
    'Set objExec = objWshShell.Exec("%comspec% /c exiftool -G " & strArg & " > " & TempFilePath)

    ' �ꎞ�t�@�C���ɁA��قǓǂ݂����ꂽexif�f�[�^����������
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tso
    Set tso = fso.OpenTextFile(TempFilePath, 2, true)
    tso.Write(objExec.StdOut.ReadAll)
    tso.Close

    Set tso = Nothing
    Set objExec = Nothing

    ' �ꎞ�t�@�C���iexif�f�[�^���������܂�Ă���j���������ŊJ��
'    objWshShell.Run("notepad.exe" & " " & TempFilePath, 1, 1)
    objWshShell.Run "notepad.exe" & " " & TempFilePath, 1, True

    ' �ꎞ�t�@�C�����폜����
    fso.DeleteFile(TempFilePath)
    Set fso = Nothing
End Sub

' �t�@�C���̃^�C���X�^���v��Exif:DateTimeOriginal�ɕύX����
Sub RestoreChangeTimeStamp(strPath)
    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    Dim strCommand : strCommand = "exiftool ""-FileModifyDate<DateTimeOriginal"" "
    ' exiftool�����s
    Dim objExec
    Set objExec = objWshShell.Exec(strCommand & strPath)
    MsgBox("���s�R�}���h : " & strCommand & strPath & vbCRLF & vbCRLF & objExec.StdOut.ReadAll)
End Sub

' Gimp�ҏW�O��Make�^�O��PENTAX��Pentax Corp.�ɏ�������
Sub ChangePentaxMakeTag(strPath)
    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    Dim strCommand : strCommand = "exiftool -if ""$Exif:Make eq 'PENTAX'"" -overwrite_original -preserve -Exif:Make=""Pentax Corp."" "
    ' exiftool�����s
    Dim objExec
    Set objExec = objWshShell.Exec(strCommand & strPath)
    MsgBox("���s�R�}���h : " & strCommand & strPath & vbCRLF & vbCRLF & objExec.StdOut.ReadAll)

    Set objExec = Nothing
    Set objWshShell = Nothing
End Sub

' Gimp�ҏW��ɔj�󂳂ꂽSoftware��ModifyDate�̏����߂��C��
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
        MsgBox("�L�����Z�����܂���")
        Exit Sub
    End If

    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    ' exiftool�����s
    Dim objExec
    Set objExec = objWshShell.Exec(strCommand & strPath)
    MsgBox("���s�R�}���h : " & strCommand & strPath & vbCRLF & vbCRLF & objExec.StdOut.ReadAll)

    Set objExec = Nothing
    Set objWshShell = Nothing
End Sub

' �t�@�C������Exif:CreateDate�ɕύX
Sub RenameFilename(strPath)

    Dim strCommand
    Dim ans
    ans = MsgBox("���s/�e�X�g�̑I��" & vbCRLF & vbCRLF & _
                "[YES] �t�@�C������ύX " & vbCRLF & _
                "[NO] �V�~�����[�V�����̂� " & vbCRLF & vbCRLF & _
                "YES or No ��I�����Ă�������" & vbCRLF & vbCRLF , vbYesNoCancel)
    If ans=vbYes Then
        strCommand = "exiftool ""-FileName<$Exif:CreateDate"" -d %Y-%m-%d-%H-%M-%S_%%f.%%e "
    ElseIf ans = vbNo Then
        strCommand = "exiftool ""-TestName<$Exif:CreateDate"" -d %Y-%m-%d-%H-%M-%S_%%f.%%e "
    Else
        MsgBox("�L�����Z�����܂���")
        WScript.Quit
    End If

    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    ' exiftool�����s
    Dim objExec
    Set objExec = objWshShell.Exec(strCommand & strPath)
    MsgBox("���s�R�}���h : " & strCommand & strPath & vbCRLF & vbCRLF & objExec.StdOut.ReadAll)

    Set objExec = Nothing
    Set objWshShell = Nothing
End Sub

' strPath���t�@�C�������ǂ����𔻒肷��i�t�@�C�����FTrue�C�f�B���N�g�����FFalse�j
Function CheckFileOnly(strPath)
    Dim fso
    CheckFileOnly = False  ' �߂�l

    ' �����ŗ^����ꂽ�t���p�X�����A�t�@�C�����ǂ����𔻒肵�A�t�@�C���̏ꍇ��True��Ԃ�
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(strPath) Then
        ' �����Ŏw�肳�ꂽ�̂̓f�B���N�g��
        CheckFileOnly = False  ' �߂�l
        Exit Function
    ElseIf fso.FileExists(strPath) <> True Then
        ' �����Ŏw�肳�ꂽ�̂̓t�@�C���ł͂Ȃ�
        CheckFileOnly = False  ' �߂�l
        Exit Function
    End If
    Set fso = Nothing

    CheckFileOnly = True  ' �߂�l
End Function

' �t�@�C���P�̂��A�t�@�C���̂���f�B���N�g���S�̉���I�����A����ɏ]�����p�X����Ԃ�
Function SelectFileOrDir(strPath)
    SelectFileOrDir = strPath ' �߂�l
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' �������f�B���N�g���̏ꍇ�́A���̂܂܏I������i�f�B���N�g������Ԃ��j
    If fso.FolderExists(strPath) Then Exit Function

    Dim ans
    ans = MsgBox("�����Ώۂ̑I��" & vbCRLF & vbCRLF & _
                "[YES] �t�@�C���P�̂�Ώ� = " & strPath & vbCRLF & _
                "[NO] �f�B���N�g�����S�̂�Ώ� = " & fso.GetParentFolderName(strPath) & vbCRLF & vbCRLF & _
                "YES or No ��I�����Ă�������" & vbCRLF & vbCRLF , vbYesNoCancel)
    If ans=vbYes Then
        SelectFileOrDir = strPath ' �߂�l
    ElseIf ans = vbNo Then
        ' �^����ꂽ�t�@�C���̐e�f�B���N�g������Ԃ�
        SelectFileOrDir = fso.GetParentFolderName(strPath) ' �߂�l
    Else
        MsgBox("�L�����Z�����܂���")
        WScript.Quit
    End If
    Set fso = Nothing
End Function
