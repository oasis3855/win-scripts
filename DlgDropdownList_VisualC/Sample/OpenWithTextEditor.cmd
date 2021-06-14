@echo off

if "%1" == "" (
    DlgDropdownList.exe "エディタで開きたいファイルを引数に指定してください"
    exit
)

DlgDropdownList.exe "起動するテキストエディタを選択してください" Notepad Terapad サクラエディタ "Visual Studio Code"
if %ERRORLEVEL% equ 1 (
    start "" notepad.exe %*
) else if %ERRORLEVEL% equ 2 (
    start "" "C:\Program Files\OnlineSoftware\TextEditor\TeraPad\TeraPad.exe" %*
) else if %ERRORLEVEL% equ 3 (
    start "" "C:\Program Files\OnlineSoftware\TextEditor\sakura\sakura.exe" %*
) else if %ERRORLEVEL% equ 4 (
    echo "end"
)
