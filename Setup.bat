@echo off
cls
:init
setlocal DisableDelayedExpansion
set cmdInvoke=1
set winSysFolder=System32
set "batchPath=%~0"
for %%k in (%0) do set batchName=%%~nk
set "vbsGetPrivileges=%temp%\OEgetPriv_%batchName%.vbs"
setlocal EnableDelayedExpansion
:checkPrivileges
NET FILE 1>NUL 2>NUL
if '%errorlevel%' == '0' ( goto gotPrivileges ) else ( goto getPrivileges )
:getPrivileges
if '%1'=='ELEV' (echo ELEV & shift /1 & goto gotPrivileges)
echo.
echo **************************************
echo �������ԱȨ���С���
echo **************************************
echo Set UAC = CreateObject^("Shell.Application"^) > "%vbsGetPrivileges%"
echo args = "ELEV " >> "%vbsGetPrivileges%"
echo For Each strArg in WScript.Arguments >> "%vbsGetPrivileges%"
echo args = args ^& strArg ^& " "  >> "%vbsGetPrivileges%"
echo Next >> "%vbsGetPrivileges%"
if '%cmdInvoke%'=='1' goto InvokeCmd 
ECHO UAC.ShellExecute "!batchPath!", args, "", "runas", 1 >> "%vbsGetPrivileges%"
goto ExecElevation
:InvokeCmd
ECHO args = "/c """ + "!batchPath!" + """ " + args >> "%vbsGetPrivileges%"
ECHO UAC.ShellExecute "%SystemRoot%\%winSysFolder%\cmd.exe", args, "", "runas", 1 >> "%vbsGetPrivileges%"
:ExecElevation
"%SystemRoot%\%winSysFolder%\WScript.exe" "%vbsGetPrivileges%" %*
exit /B
:gotPrivileges
setlocal & cd /d %~dp0
if '%1'=='ELEV' (del "%vbsGetPrivileges%" 1>nul 2>nul  &  shift /1)


@echo off
color 06
mode con lines=30 cols=65

title MSO_Setup_Helper
echo     MSO_Setup_Helper v1.0.1
echo.
echo     ���������ʼ���� [Ctrl+C] �˳�����
pause >nul

:menu
title �˵���ѡ����
cls
echo ---------------------------------------------------------------
echo     ��ѡ������
echo ---------------------------------------------------------------
echo     [A]����
echo     [B]��Ϣ
echo ---------------------------------------------------------------
echo     [Ctrl+C] �˳�
echo ---------------------------------------------------------------
echo.
:menucho
set choice=
set /p choice=�������Ӧ����ĸ���� [Enter] ��ʼ:
if not "%Choice%"=="" SET Choice=%Choice:~0,1%
if /i "%choice%"=="a" goto menusetup
if /i "%choice%"=="b" goto information
echo ѡ����Ч������������
echo.
goto menucho
pause >nul
cls

:menusetup
title ����ѡ����
cls
echo ---------------------------------------------------------------
echo     ��ѡ������
echo ---------------------------------------------------------------
echo     [A]���� Office 
echo ---------------------------------------------------------------
echo     [Z] �����ϼ�
echo     [Ctrl+C] �˳�
echo ---------------------------------------------------------------
echo.
:menusetupcho
set choice=
set /p choice=�������Ӧ����ĸ���� [Enter] ��ʼ:
if not "%Choice%"=="" SET Choice=%Choice:~0,1%
if /i "%choice%"=="a" goto menuoffice
if /i "%choice%"=="z" goto menu
echo ѡ����Ч������������
echo.
goto menusetupcho
pause >nul
cls

:menuoffice
cls
title ���� Office 365 / 2019��ѡ����
echo ---------------------------------------------------------------
echo     ��ѡ������
echo ---------------------------------------------------------------
echo     [A]���߰�װ Office 365 32 λ
echo     [B]���߰�װ Office 365 64 λ
echo     [C]���߰�װ Office 2019 32 λ
echo     [D]���߰�װ Office 2019 64 λ
echo     [E]���� Office 365 32 λ��װ��
echo     [F]���� Office 365 64 λ��װ��
echo     [G]���� Office 2019 32 λ��װ��
echo     [H]���� Office 2019 64 λ��װ��
echo     [I]���� Office ����ͨ��˵��
echo     [J]ʹ�ùٷ������Ƴ� Office
echo     [K]��װ Office ���Ը����� 32 λ
echo     [L]��װ Office ���Ը����� 64 λ
echo ---------------------------------------------------------------
echo     [Z] �����ϼ�
echo     [Ctrl+C] �˳�
echo ---------------------------------------------------------------
echo.
:menuofficecho
set choice=
set /p choice=�������Ӧ����ĸ���� [Enter] ��ʼ:
if not "%Choice%"=="" SET Choice=%Choice:~0,1%
if /i "%choice%"=="a" goto start1
if /i "%choice%"=="b" goto start2
if /i "%choice%"=="c" goto start3
if /i "%choice%"=="d" goto start4
if /i "%choice%"=="e" goto start5
if /i "%choice%"=="f" goto start6
if /i "%choice%"=="g" goto start7
if /i "%choice%"=="h" goto start8
if /i "%choice%"=="i" goto start9
if /i "%choice%"=="j" goto start10
if /i "%choice%"=="k" goto start11
if /i "%choice%"=="l" goto start12
if /i "%choice%"=="z" goto menusetup
echo ѡ����Ч������������
echo.
goto menuofficecho
pause >nul
cls

:start1
cls
title ���߰�װOffice 365 32 λ
cd /d .\files\
cmd /k "setup.exe /configure configuration-Office365-x86.xml"
title ���!
echo ���!
echo. & pause
goto menuoffice

:start2
cls
title ���߰�װOffice 365 64 λ
cd /d .\files\
cmd /k "setup.exe /configure configuration-Office365-x64.xml"
title ���!
echo ���!
echo. & pause
goto menuoffice

:start3
cls
title ���߰�װOffice 2019 32 λ
cd /d .\files\
cmd /k "setup.exe /configure configuration-Office2019Enterprise-x86.xml"
title ���!
echo ���!
echo. & pause
goto menuoffice

:start4
cls
title ���߰�װOffice 2019 64 λ
cd /d .\files\
cmd /k "setup.exe /configure configuration-Office2019Enterprise-x64.xml
title ���!
echo ���!
echo. & pause
goto menuoffice

:start5
cls
title ����Office 365 32 λ
cd /d .\files\
cmd /k "setup.exe /download configuration-Office365-x86.xml"
title ���!
echo ���!
echo. & pause
goto menuoffice

:start6
cls
title ����Office 365 64 λ
cd /d .\files\
cmd /k "setup.exe /download configuration-Office365-x64.xml"
title ���!
echo ���!
echo. & pause
goto menuoffice

:start7
cls
title ����Office 2019 32 λ
cd /d .\files\
cmd /k "setup.exe /download configuration-Office2019Enterprise-x86.xml"
title ���!
echo ���!
echo. & pause
goto menuoffice

:start8
cls
title ����Office 2019 64 λ
cd /d .\files\
cmd /k "setup.exe /download configuration-Office2019Enterprise-x64.xml"
title ���!
echo ���!
echo. & pause
goto menuoffice

:start9
cls
title ���� Office ����ͨ��
echo     ���� Office ����ͨ����
echo     ʹ�ñ��ű���װĬ��Ϊÿ��ͨ��
echo     ������ģ������files�ļ��б༭��Ӧ�汾xml�ļ��еĸ���ͨ��
echo     �ļ��л��ҵ� Channel="Monthly"
echo     ���� Monthly �����޸�Ϊ���桾���е�ͨ��
echo ---------------------------------------------------------------
echo     �����г�����ͨ�����ƣ�
echo     Office 2019 ��ҵ���ڰ桾PerpetualVL2019��
echo     ����ͨ����Broad��
echo     ����ͨ�������򣩡�Targeted��
echo     ÿ��ͨ����Monthly��
echo     ÿ��ͨ�������򣩡�Insiders��
echo     ����ͨ����Ԥ������ƻ�����InsiderFast��
echo     ����ͨ�����ڲ����ԣ���Dogfood��
echo ---------------------------------------------------------------
echo     ����������ϼ��˵���[Ctrl+C]�ر�
echo ---------------------------------------------------------------
echo. & pause
goto menuoffice

:start10
cls
title ʹ�ùٷ������Ƴ�Office
echo ���ڵ���������ɲ���
cd /d .\files\clean\
cmd /k "o15-ctrremove.diagcab"
title ���!
echo ���!
echo. & pause
goto menuoffice

:start11
cls
title ��װOffice���Ը����� 32λ
echo ��װOffice���Ը����� 32λ
echo. & pause
echo ��ȴ�����������ɲ���
cd /d .\files\
cmd /k "setuplanguagepack.x86.zh-cn_.exe"
title ���!
echo ���!
echo. & pause
goto menuoffice

:start12
cls
title ��װOffice���Ը����� 64λ
echo ��װOffice���Ը����� 64λ
echo. & pause
echo ��ȴ�����������ɲ���
cd /d .\files\
cmd /k "setuplanguagepack.x64.zh-cn_.exe"
title ���!
echo ���!
echo. & pause
goto menuoffice

:information
cls
title ��Ϣ
echo ---------------------------------------------------------------
echo     MSO_Setup_Helper
echo     �������ڣ�2021/6/23
echo     �汾��v1.0.1
echo     (c) zhxy-CN, Released under the MIT License.
echo ---------------------------------------------------------------
echo     [Z] �����ϼ�
echo     [Ctrl+C]�˳�
echo ---------------------------------------------------------------
echo.
:informationcho
set choice=
set /p choice=�������Ӧ����ĸ���� [Enter] ��ʼ:
if not "%Choice%"=="" SET Choice=%Choice:~0,1%
if /i "%choice%"=="z" goto menu
echo ѡ����Ч������������
echo.
goto informationcho
pause >nul
cls
