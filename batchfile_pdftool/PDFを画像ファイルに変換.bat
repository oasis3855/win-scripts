@echo off
:: **********
:: PDF���摜�t�@�C���ɕϊ�(imagemagick).bat
:: 
:: imagemagick ���s�t�@�C���̂���f�B���N�g���ɁAPATH��ʂ��Ă����K�v������܂�
:: ���ϐ� MAGICK_CODER_MODULE_PATH �� MAGICK_HOME ��ݒ肷��K�v������
:: 
:: Version 1.0 (2022/03/21)
:: **********

echo PDF�̑S�y�[�W��jpg/png/bmp���ɕϊ����܂�

:: �����`�F�b�N
if [%1]==[] (
    echo usage info : �����ɏ����Ώ�PDF�t�@�C�����w�肵�Ă�������
    exit /b
)
if not exist %1 (
    echo error : �t�@�C�� %1 ��������܂���
    exit /b
)

set OUTPUT_DIR=pdf2jpg
if not exist %OUTPUT_DIR%\ (
    mkdir %OUTPUT_DIR%
    echo  �V�K�쐬�����f�B���N�g�� %OUTPUT_DIR%\ �ɏo�͂��܂�
) else (
    echo  �����̃f�B���N�g�� %OUTPUT_DIR%\ �ɏo�͂��܂�
    echo  ^(�����̃t�@�C�������݂���ꍇ�́A�㏑������܂���^)
)
echo;

:: ���\��
echo  �����Ώۃt�@�C�� : %1
echo  �o��jpg�t�@�C�� : %OUTPUT_DIR%\001.ext �` %OUTPUT_DIR%\999.ext
echo;

:: ���[�U���́i�o�͊g���q�j
set USR_EXT=jpg
set /P USR_EXT="�o�̓t�@�C���g���q�����. �f�t�H���g�l %USR_EXT% (jpg, png, bmp, tiff) : "

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
    echo error : �w��ȊO�̊g���q�����͂���܂���
    pause
    exit /b
)
echo  �g���q %USR_EXT% ���I������܂���
echo  �o�̓t�@�C���� 001.%USR_EXT% , 002.%USR_EXT% , ... 999.%USR_EXT% �ƂȂ�܂�

:: ���[�U���́i�𑜓x DPI�j
set USR_DPI=200
set /P USR_DPI="DPI�𐮐��œ���. �f�t�H���g�l %USR_DPI% (72-800) : "

if %USR_DPI% lss 72 (
   echo DPI��72�ȏ����͂��Ă�������
   pause
   exit /b
)
if %USR_DPI% gtr 800 (
   echo DPI��800�ȉ�����͂��Ă�������
   pause
   exit /b
)
echo  DPI�����͂���܂��� : -density %USR_DPI%
echo;


:: ���[�U���́ijpg�ۑ��N�I���e�B�j
set USR_QUALITY=75
set /P USR_QUALITY="jpeg�ۑ��N�I���e�B�𐮐��œ���. �f�t�H���g�l %USR_QUALITY% (1-100) : "

if %USR_QUALITY% lss 1 (
   echo jpeg�ۑ��N�I���e�B��1�ȏ����͂��Ă�������
   pause
   exit /b
)
if %USR_QUALITY% gtr 100 (
   echo jpeg�ۑ��N�I���e�B��100�ȉ�����͂��Ă�������
   pause
   exit /b
)
echo  Quality�����͂���܂��� : -quality %USR_QUALITY%
echo;

:: ���s�O�̃R�}���h�\��
echo �R�}���h
echo   magick -density %USR_DPI% -quality %USR_QUALITY% %1.pdf %OUTPUT_DIR%\%%03d.%USR_EXT%
echo �����s���܂�

pause

:: imagemagick �R�}���h�����s
magick -density %USR_DPI% -quality %USR_QUALITY% %1 %OUTPUT_DIR%\%%03d.%USR_EXT%

:: �I�����b�Z�[�W
echo imagemagick �R�}���h�����I�� (�R�}���h�߂�l %ERRORLEVEL%)
pause > nul

