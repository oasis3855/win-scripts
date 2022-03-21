@echo off
:: **********
:: PDFを画像ファイルに変換(imagemagick).bat
:: 
:: imagemagick 実行ファイルのあるディレクトリに、PATHを通しておく必要があります
:: 環境変数 MAGICK_CODER_MODULE_PATH と MAGICK_HOME を設定する必要がある
:: 
:: Version 1.0 (2022/03/21)
:: **********

echo PDFの全ページをjpg/png/bmp等に変換します

:: 引数チェック
if [%1]==[] (
    echo usage info : 引数に処理対象PDFファイルを指定してください
    exit /b
)
if not exist %1 (
    echo error : ファイル %1 が見つかりません
    exit /b
)

set OUTPUT_DIR=pdf2jpg
if not exist %OUTPUT_DIR%\ (
    mkdir %OUTPUT_DIR%
    echo  新規作成したディレクトリ %OUTPUT_DIR%\ に出力します
) else (
    echo  既存のディレクトリ %OUTPUT_DIR%\ に出力します
    echo  ^(同名のファイルが存在する場合は、上書きされません^)
)
echo;

:: 情報表示
echo  処理対象ファイル : %1
echo  出力jpgファイル : %OUTPUT_DIR%\001.ext ～ %OUTPUT_DIR%\999.ext
echo;

:: ユーザ入力（出力拡張子）
set USR_EXT=jpg
set /P USR_EXT="出力ファイル拡張子を入力. デフォルト値 %USR_EXT% (jpg, png, bmp, tiff) : "

if [%USR_EXT%]==[] (
    set USR_EXT=jpg
) else if "%USR_EXT%"== "jpg" (
    echo;
) else if "%USR_EXT%"== "png" (
    echo;
) else if "%USR_EXT%"== "bmp" (
    echo;
) else if "%USR_EXT%"== "tiff" (
    echo;
) else (
    echo error : 指定以外の拡張子が入力されました
    pause
    exit /b
)
echo  拡張子 %USR_EXT% が選択されました
echo  出力ファイルは 001.%USR_EXT% , 002.%USR_EXT% , ... 999.%USR_EXT% となります

:: ユーザ入力（解像度 DPI）
set USR_DPI=200
set /P USR_DPI="DPIを整数で入力. デフォルト値 %USR_DPI% (72-800) : "

if %USR_DPI% lss 72 (
   echo DPIは72以上を入力してください
   pause
   exit /b
)
if %USR_DPI% gtr 800 (
   echo DPIは800以下を入力してください
   pause
   exit /b
)
echo  DPIが入力されました : -density %USR_DPI%
echo;


:: ユーザ入力（jpg保存クオリティ）
set USR_QUALITY=75
set /P USR_QUALITY="jpeg保存クオリティを整数で入力. デフォルト値 %USR_QUALITY% (1-100) : "

if %USR_QUALITY% lss 1 (
   echo jpeg保存クオリティは1以上を入力してください
   pause
   exit /b
)
if %USR_QUALITY% gtr 100 (
   echo jpeg保存クオリティは100以下を入力してください
   pause
   exit /b
)
echo  Qualityが入力されました : -quality %USR_QUALITY%
echo;

:: 実行前のコマンド表示
echo コマンド
echo   magick -density %USR_DPI% -quality %USR_QUALITY% %1.pdf %OUTPUT_DIR%\%%03d.%USR_EXT%
echo を実行します

pause

:: imagemagick コマンドを実行
magick -density %USR_DPI% -quality %USR_QUALITY% %1 %OUTPUT_DIR%\%%03d.%USR_EXT%

:: 終了メッセージ
echo imagemagick コマンド処理終了 (コマンド戻り値 %ERRORLEVEL%)
pause > nul

