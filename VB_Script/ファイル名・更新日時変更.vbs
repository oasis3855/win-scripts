Option Explicit

' **********************************
' �t�@�C�����E�X�V�����ύX.vbs
'
' ���̃t�@�C���� Shift JIS �ŕۑ����邱��
' �ʓr�K�v�ȃv���O����
'   DlgDropdownList �i�X�N���v�g����Ăяo���_�C�A���O�{�b�N�X�P�̕\���j
'   https://github.com/oasis3855/WindowsScripts/tree/main/DlgDropdownList_VisualC
' **********************************

Call Main()

Sub Main()

    Dim modeFolder      ' �f�B���N�g�����S�t�@�C���ΏۂƂ��郂�[�h�̏ꍇ�� True
    Dim strTargetPath   ' �Ώۃt�@�C�� �܂��� �f�B���N�g��

    ' ���̃X�N���v�g�̈�����1���ǂ����̃`�F�b�N
    if Wscript.Arguments.Count <> 1 then
        MsgBox("�����i�ΏۂƂ���t�@�C���j��1�w�肵�Ă�������")
        WScript.Quit
    End If

    ' �����Ŏw�肳�ꂽ���̂��A�t�@�C�����A�f�B���N�g�����`�F�b�N����
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim strArg
    strArg = Wscript.Arguments(0)
    If fso.FolderExists(strArg) Then
        ' �����̓f�B���N�g��
        modeFolder = True
        strTargetPath = strArg
    ElseIf fso.FileExists(strArg) Then
        ' �����̓t�@�C��
        modeFolder = False
    Else
        MsgBox(strArg & " �͑��݂��܂���")
        WScript.Quit
    End If

    ' �t�@�C���������Ŏw�肳��Ă���Ƃ��A�t�@�C���P�̂��A�t�@�C���̂���f�B���N�g���S�̂��̑I��������
    Dim ans
    If modeFolder = False Then
        ans = MsgBox("�����Ώۂ̑I��" & vbCRLF & vbCRLF & _
                    "[YES] �t�@�C���P�̂�Ώ� = " & strArg & vbCRLF & _
                    "[NO] �f�B���N�g�����S�̂�Ώ� = " & fso.GetParentFolderName(strArg) & vbCRLF & vbCRLF & _
                    "YES or No ��I�����Ă�������" & vbCRLF & vbCRLF , vbYesNoCancel)
        If ans=vbYes Then
            strTargetPath = strArg
        ElseIf ans = vbNo Then
            modeFolder = True
            strTargetPath = fso.GetParentFolderName(strArg)
        Else
            MsgBox("�L�����Z�����܂���")
            WScript.Quit
        End If
    End If
    set fso = Nothing

    ' ���X�g�{�b�N�X�̃_�C�A���O�ŁA�@�\�̑I�����s��
    Dim numUserSelect
    numUserSelect = SelectTaskDialog(strTargetPath, modeFolder)

    ' �t�@�C�����̕ύX�̏ꍇ�A���X�g�{�b�N�X�̃_�C�A���O�ŁA�Ώۂ��x�[�X�l�[�����g���q����I������
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
            MsgBox("���[�U�ɂ���ăL�����Z������܂���")
            Exit Sub
        End If
    End If

    Dim counter     ' �������������t�@�C���̌��i���ʕ\���p�j
    counter = 0
    If numUserSelect = 1 then
        counter = ChangeFnames(strTargetPath, modeFolder, "small", modeBasenameExt)
        MsgBox(counter & " �̃t�@�C������ύX���܂���")
    ElseIf numUserSelect = 2 then
        counter = ChangeFnames(strTargetPath, modeFolder, "LARGE", modeBasenameExt)
        MsgBox(counter & " �̃t�@�C������ύX���܂���")
    ElseIf numUserSelect = 3 then
        counter = TouchFiles(strTargetPath, modeFolder, Now)
        MsgBox(counter & " �̃t�@�C���̍X�V���������ݓ����ɕύX���܂���")
    ElseIf numUserSelect = 4 then
        Dim SelectedDateTime
        SelectedDateTime = InputBox("�X�V�����̃��[�U����", "�Ώۃt�@�C��" & strTargetPath, Now)
        counter = TouchFiles(strTargetPath, modeFolder, SelectedDateTime)
        MsgBox(counter & " �̃t�@�C���̍X�V������ " & SelectedDateTime & " �ɕύX���܂���")
    Else
        MsgBox("���[�U�ɂ���ăL�����Z������܂���")
    End If

End Sub

' �@�\�I���_�C�A���O�i���X�g�{�b�N�X���I���j��\������
Function SelectTaskDialog(strTargetPath, modeFolder)
    SelectTaskDialog = 0    ' �f�t�H���g�̖߂�l

    Dim strMsg
    If modeFolder = True Then strMsg = "[�f�B���N�g��]  " Else strMsg = "[�t�@�C���P��]  "

    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    SelectTaskDialog = objWshShell.Run("DlgDropdownList.exe" & " " & """" & strMsg & strTargetPath & "  �ɑ΂��čs��������I�����Ă�������"" " & _
                        "�t�@�C�������������ɕύX �t�@�C������啶���ɕύX �ŏI�X�V�������ݓ����ɕύX �ŏI�X�V�����w�肵�������ɕύX", 1, True)
End Function

' ���l�[���ΏۑI���_�C�A���O�i���X�g�{�b�N�X���I���j��\������
Function SelectBaseExtDialog(strTargetPath, modeLargeSmall)
    SelectBaseExtDialog = 0    ' �f�t�H���g�̖߂�l

    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    Dim strLargeSmall
    If modeLargeSmall = "small" Then
        strLargeSmall = "��������"
    Else
        strLargeSmall = "�啶����"
    End If
    Dim retNum
    SelectBaseExtDialog = objWshShell.Run("DlgDropdownList.exe" & " " & """" & strTargetPath & " �̃t�@�C������" & strLargeSmall & "�̑Ώۂ�I�����Ă�������"" " & _
                        "�x�[�X�l�[���̂ݑΏ� �g���q�̂ݑΏ� �x�[�X�l�[���Ɗg���q�̗�����Ώ�", 1, True)
End Function

' �w�肵���t�@�C���i�܂��̓f�B���N�g�����̑S�t�@�C���j�̍X�V������ύX����
Function TouchFiles(strTargetPath, modeFolder, SelectedDateTime)
    ' �߂�l�̏����ݒ�i���������������t�@�C�����j
    TouchFiles = 0

    ' �P�̃t�@�C���̂Ƃ�
    If modeFolder = False Then
        If TouchFile(strTargetPath, SelectedDateTime) Then
            ' �t�@�C��1�̏����������������̖߂�l
            TouchFiles = 1
        End If
        Exit Function
    End If

    ' ����������̃t�H�[�}�b�g�`�F�b�N
    If IsDate(SelectedDateTime) = False Then
        TouchFiles = 0
        MsgBox("�����̏������K��O�ł� : " & SelectedDateTime)
        exit Function
    End If

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim objShell
    Set objShell = CreateObject("Shell.Application")

    Dim objFiles
    Dim objFile
    ' �w�肳�ꂽ�f�B���N�g�����̑S�t�@�C���ꗗ
    Set objFiles = fso.GetFolder(strTargetPath).files
    ' �t�@�C��1������
    For Each objFile In objFiles
        If TouchFile(objFile, SelectedDateTime) Then
            ' �t�@�C��1���̏����������������A�߂�l�i�����t�@�C�����J�E���^�j���C���N�������g����
            TouchFiles = TouchFiles + 1
        End If
    Next

    set objFile = Nothing
    set objFiles = Nothing

End Function

' �w�肵���t�@�C���P�̂̍X�V������ύX����
Function TouchFile(strTargetPath, SelectedDateTime)
    ' �߂�l�̏����ݒ�i��������True��Ԃ��j
    TouchFile = True

    ' ����������̃t�H�[�}�b�g�`�F�b�N
    If IsDate(SelectedDateTime) = False Then
        TouchFile = False
        MsgBox("�����̏������K��O�ł� : " & SelectedDateTime)
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

    ' FolderItem ������������ꂽ�ꍇ�ɁA�X�V������ύX����
    If (Not objFolder.Items.Item(fso.GetFileName(strTargetPath)) is Nothing) then
        objFolder.Items.Item(fso.GetFileName(strTargetPath)).ModifyDate = SelectedDateTime
    Else
        TouchFile = False
    End If

    set objFolder = Nothing

End Function

' �w�肵���t�@�C���i�܂��̓f�B���N�g�����̑S�t�@�C���j�̃t�@�C�������A�啶�����E������������
Function ChangeFnames(strTargetPath, modeFolder, modeLargeSmall, modeBasenameExt)
    ' �߂�l�̏����ݒ�i���������������t�@�C�����j
    ChangeFnames = 0
    ' �P�̃t�@�C���̎�
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
    ' �w�肳�ꂽ�f�B���N�g�����̑S�t�@�C���ꗗ
    Set objFiles = fso.GetFolder(strTargetPath).files
    ' �t�@�C��1������
    For Each objFile In objFiles
        If ChangeFname(objFile, modeLargeSmall, modeBasenameExt) = True then
            ChangeFnames = ChangeFnames + 1
        End If
    Next

    set objFile = Nothing
    set objFiles = Nothing



End Function

' �w�肵���t�@�C���P�̂̃t�@�C�������A�啶�����E������������
Function ChangeFname(strTargetPath, modeLargeSmall, modeBasenameExt)
    ' �߂�l�̏����ݒ�i��������True��Ԃ��j
    ChangeFname = False

    ' Scripting.FileSystemObject �� File.Name �ł̃t�@�C�����ύX�́A�啶����������ʂȂ����߃G���[�ƂȂ�
    ' ���̂��߁AShell.Application �� FolderItem.Name �Ńt�@�C�����ύX����

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim objShell
    Set objShell = CreateObject("Shell.Application")

    ' �Ώۃt�@�C�������݂��邱�Ƃ��m�F
    If fso.FileExists(strTargetPath) <> True Then
        ChangeFname = False
        MsgBox("�Ώۃt�@�C����������Ȃ� : " & strTargetPath)
        Exit Function
    End If

    ' �p�X���𕪉����A�f�B���N�g���A�t�@�C���{�́A�g���q�̕�����Ɋi�[
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
    ' �ύX��̃t�@�C�����̍쐬
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

    ' �t�@�C�����ύX��ƁA���ꖼ�̃t�@�C���i�啶�����������ʁj�A���ꖼ�t�H���_�����݂��Ȃ����Ƃ��m�F
    If IsFileExistCaseSensitive(dirname & "\" & filenameNew) = True Or fso.FolderExists(dirname & "\" & filenameNew) = True Then
        ChangeFname = False
        ' �f�o�b�O���̃��b�Z�[�W
        ' MsgBox("���ꖼ�̃t�@�C�������� : " & filenameNew)
        Exit Function
    End If
    ' �t�@�C�����̕ύX
    objFolderitem.Name = filenameNew
    ' �����������̖߂�l
    ChangeFname = True
End Function

' �t�@�C�����̑啶������������ʂ��āA�t�@�C�������݂��邩�`�F�b�N����
Function IsFileExistCaseSensitive(strTargetPath)
    ' �߂�l�̏����ݒ�
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
