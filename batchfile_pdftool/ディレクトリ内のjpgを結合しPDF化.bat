@echo off
:: **********
:: �f�B���N�g������jpg��������PDF��(imagemagick).bat
:: 
:: imagemagick ���s�t�@�C���̂���f�B���N�g���ɁAPATH��ʂ��Ă����K�v������܂�
:: ���ϐ� MAGICK_CODER_MODULE_PATH �� MAGICK_HOME ��ݒ肷��K�v������
:: 
:: Version 1.0 (2022/03/21)
:: **********

echo �f�B���N�g������jpg��������PDF��(imagemagick)
echo;

:: �����`�F�b�N
if [%1]==[] (
    echo usage info : �����ɏ����Ώۃf�B���N�g�����̃t�@�C�����w�肵�Ă�������
    exit /b
)
if not exist %1 (
    echo error : �t�@�C�� %1 ��������܂���
    exit /b
)


:: ���݃f�B���N�g���𓾂�i%~d1 : �J�����g�h���C�u + %~p1 : �J�����g�f�B���N�g���j
set FIILE_DIR=
set FILE_DIR=%~dp1
echo ���̓f�B���N�g�� : %FILE_DIR%
echo �o�̓t�@�C�� : %FILE_DIR%output.pdf
echo;

:: �o�̓t�@�C���㏑���`�F�b�N
if exist "%FILE_DIR%output.pdf" (
    echo  ^(�o�̓t�@�C���Ɠ����̃t�@�C��������. �㏑�����܂�^)
    echo;
)

:: ���s�O�̃R�}���h�\��
echo �R�}���h
echo   magick %FILE_DIR%*.jpg %FILE_DIR%output.pdf
echo �����s���܂�
echo;
pause

:: imagemagick �R�}���h�����s
magick "%FILE_DIR%*.jpg" "%FILE_DIR%output.pdf"

:: �I�����b�Z�[�W
echo imagemagick �R�}���h�����I�� (�R�}���h�߂�l %ERRORLEVEL%)
pause > nul

