Option Explicit

Dim strArg
Dim strTargetPath
Dim fso
Dim objWshShell

' �����������ꍇ
if Wscript.Arguments.Count < 1 then
    MsgBox("�G�f�B�^�ŊJ�������t�@�C���������Ɏw�肵�Ă�������")
    WScript.Quit
End If

' �v���O�����̈����iExplorer�̑���(Send To)�ɂ��A�t���p�X�̃t�@�C�������i�[����Ă���
strArg = Wscript.Arguments(0)

' ����1�ŗ^����ꂽ���̂��A�t�@�C�����ǂ����𔻒肵�A�t�@�C���ȊO�̏ꍇ�͏I������
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FolderExists(strArg) Then
    MsgBox("�����Ŏw�肳�ꂽ�̂̓f�B���N�g���ł��B�����ΏۊO�̂��ߏI�����܂�")
    WScript.Quit
ElseIf fso.FileExists(strArg) <> True Then
    MsgBox("�����Ŏw�肳�ꂽ�̂̓t�@�C���ł͂���܂���B�����ΏۊO�̂��ߏI�����܂�")
    WScript.Quit
End If


Set objWshShell = CreateObject("WScript.Shell")

' �G�f�B�^�̎�ނ�I������_�C�A���O�{�b�N�X��\�����A�I�������l��ExitCode�Ƃ��ē���
' Run���\�b�h�̑�2�����͎��s����E�C���h�E�̏�ԁA��3�����̓v���Z�X�̏C����҂��ǂ����̐ݒ�
Dim numExitCode
numExitCode = objWshShell.Run("DlgDropdownList.exe" & " " & "�N������e�L�X�g�G�f�B�^��I�����Ă������� Notepad Terapad �T�N���G�f�B�^ ""Visual Studio Code""", 1, True)

' �I�������G�f�B�^�ŁA�w�肳�ꂽ�t�@�C���istrArg�j���J��
' Exec���\�b�h�́A�v���Z�X�̏I����҂����ɐ����Ԃ�
If numExitCode = 1 Then
    objWshShell.Exec("notepad.exe" & " " & strArg)
ElseIf numExitCode = 2 Then
    objWshShell.Exec("C:\Program Files\OnlineSoftware\TextEditor\TeraPad\TeraPad.exe" & " " & strArg)
ElseIf numExitCode = 3 Then
    objWshShell.Exec("C:\Program Files\OnlineSoftware\TextEditor\sakura\sakura.exe" & " " & strArg)
ElseIf numExitCode = 4 Then
    objWshShell.Exec("C:\Program Files\Microsoft VS Code\Code.exe" & " " & strArg)
End If
