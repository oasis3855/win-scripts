@echo off

if "%1" == "" (
    DlgDropdownList.exe "�G�f�B�^�ŊJ�������t�@�C���������Ɏw�肵�Ă�������"
    exit
)

DlgDropdownList.exe "�N������e�L�X�g�G�f�B�^��I�����Ă�������" Notepad Terapad �T�N���G�f�B�^ "Visual Studio Code"
if %ERRORLEVEL% equ 1 (
    start "" notepad.exe %*
) else if %ERRORLEVEL% equ 2 (
    start "" "C:\Program Files\OnlineSoftware\TextEditor\TeraPad\TeraPad.exe" %*
) else if %ERRORLEVEL% equ 3 (
    start "" "C:\Program Files\OnlineSoftware\TextEditor\sakura\sakura.exe" %*
) else if %ERRORLEVEL% equ 4 (
    echo "end"
)
