@echo off
:: **********
:: PDF縮小(Ghostscript).bat
:: 
:: Version 1.0 (2022/03/21)
:: **********

echo PDF内の画像圧縮率を変更し保存します

:: 引数チェック
if [%1]==[] (
    echo usage info : 引数に処理対象PDFファイルを指定してください
    exit /b
)
if not exist %1 (
    echo error : ファイル %1 が見つかりません
    exit /b
)

:: 情報表示
echo  処理対象ファイル : %1
echo  処理後保存ファイル : %1._gs.pdf
if exist %1._gs.pdf (
    echo;
    echo  ^(ファイルは上書きされます^)
)
echo;


:: ユーザ入力（圧縮率）
set USR_RESOLUTION=1
set /P USR_RESOLUTION="圧縮率を入力 screen:1, ebook:2, printer:3, prepress:4 [1-4] : "

if %USR_RESOLUTION% lss 1 (
   echo 1 以上を入力してください
   pause
   exit /b
)
if %USR_RESOLUTION% gtr 4 (
   echo 4 以下を入力してください
   pause
   exit /b
)

set SW_RESOLUTION=default
if %USR_RESOLUTION% == 1 set SW_RESOLUTION=screen
if %USR_RESOLUTION% == 2 set SW_RESOLUTION=ebook
if %USR_RESOLUTION% == 3 set SW_RESOLUTION=printer
if %USR_RESOLUTION% == 4 set SW_RESOLUTION=prepress

echo 圧縮率 %SW_RESOLUTION% を適用します
echo;

:: 実行前のコマンド表示
echo コマンド
echo   gswin64.exe -dNOPAUSE -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/%SW_RESOLUTION% -sOutputFile=%1._gs.pdf %1
echo を実行します
echo;
echo;
echo GhostScriptの画面を終了するときは quit と入力してください
echo;
pause

:: Ghostscript コマンド実行
gswin64.exe -dNOPAUSE -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/%SW_RESOLUTION% -sOutputFile=%1._gs.pdf %1


