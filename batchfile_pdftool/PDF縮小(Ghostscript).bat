@echo off
:: **********
:: PDF�k��(Ghostscript).bat
:: 
:: Version 1.0 (2022/03/21)
:: **********

echo PDF���̉摜���k����ύX���ۑ����܂�

:: �����`�F�b�N
if [%1]==[] (
    echo usage info : �����ɏ����Ώ�PDF�t�@�C�����w�肵�Ă�������
    exit /b
)
if not exist %1 (
    echo error : �t�@�C�� %1 ��������܂���
    exit /b
)

:: ���\��
echo  �����Ώۃt�@�C�� : %1
echo  ������ۑ��t�@�C�� : %1._gs.pdf
if exist %1._gs.pdf (
    echo;
    echo  ^(�t�@�C���͏㏑������܂�^)
)
echo;


:: ���[�U���́i���k���j
set USR_RESOLUTION=1
set /P USR_RESOLUTION="���k������� screen:1, ebook:2, printer:3, prepress:4 [1-4] : "

if %USR_RESOLUTION% lss 1 (
   echo 1 �ȏ����͂��Ă�������
   pause
   exit /b
)
if %USR_RESOLUTION% gtr 4 (
   echo 4 �ȉ�����͂��Ă�������
   pause
   exit /b
)

set SW_RESOLUTION=default
if %USR_RESOLUTION% == 1 set SW_RESOLUTION=screen
if %USR_RESOLUTION% == 2 set SW_RESOLUTION=ebook
if %USR_RESOLUTION% == 3 set SW_RESOLUTION=printer
if %USR_RESOLUTION% == 4 set SW_RESOLUTION=prepress

echo ���k�� %SW_RESOLUTION% ��K�p���܂�
echo;

:: ���s�O�̃R�}���h�\��
echo �R�}���h
echo   gswin64.exe -dNOPAUSE -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/%SW_RESOLUTION% -sOutputFile=%1._gs.pdf %1
echo �����s���܂�
echo;
echo;
echo GhostScript�̉�ʂ��I������Ƃ��� quit �Ɠ��͂��Ă�������
echo;
pause

:: Ghostscript �R�}���h���s
gswin64.exe -dNOPAUSE -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/%SW_RESOLUTION% -sOutputFile=%1._gs.pdf %1


