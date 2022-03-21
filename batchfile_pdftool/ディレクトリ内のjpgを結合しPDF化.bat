@echo off
:: **********
:: ディレクトリ内のjpgを結合しPDF化(imagemagick).bat
:: 
:: imagemagick 実行ファイルのあるディレクトリに、PATHを通しておく必要があります
:: 環境変数 MAGICK_CODER_MODULE_PATH と MAGICK_HOME を設定する必要がある
:: 
:: Version 1.0 (2022/03/21)
:: **********

echo ディレクトリ内のjpgを結合しPDF化(imagemagick)
echo;

:: 引数チェック
if [%1]==[] (
    echo usage info : 引数に処理対象ディレクトリ内のファイルを指定してください
    exit /b
)
if not exist %1 (
    echo error : ファイル %1 が見つかりません
    exit /b
)


:: 現在ディレクトリを得る（%~d1 : カレントドライブ + %~p1 : カレントディレクトリ）
set FIILE_DIR=
set FILE_DIR=%~dp1
echo 入力ディレクトリ : %FILE_DIR%
echo 出力ファイル : %FILE_DIR%output.pdf
echo;

:: 出力ファイル上書きチェック
if exist "%FILE_DIR%output.pdf" (
    echo  ^(出力ファイルと同名のファイルが存在. 上書きします^)
    echo;
)

:: 実行前のコマンド表示
echo コマンド
echo   magick %FILE_DIR%*.jpg %FILE_DIR%output.pdf
echo を実行します
echo;
pause

:: imagemagick コマンドを実行
magick "%FILE_DIR%*.jpg" "%FILE_DIR%output.pdf"

:: 終了メッセージ
echo imagemagick コマンド処理終了 (コマンド戻り値 %ERRORLEVEL%)
pause > nul

