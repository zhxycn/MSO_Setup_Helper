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
echo 请求管理员权限中……
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
echo     按任意键开始，或按 [Ctrl+C] 退出程序！
pause >nul

:menu
title 菜单－选择功能
cls
echo ---------------------------------------------------------------
echo     请选择任务。
echo ---------------------------------------------------------------
echo     [A]部署。
echo     [B]信息。
echo ---------------------------------------------------------------
echo     [Ctrl+C] 退出。
echo ---------------------------------------------------------------
echo.
:menucho
set choice=
set /p choice=请输入对应的字母，按 [Enter] 开始:
if not "%Choice%"=="" SET Choice=%Choice:~0,1%
if /i "%choice%"=="a" goto menusetup
if /i "%choice%"=="b" goto information
echo 选择无效，请重新输入
echo.
goto menucho
pause >nul
cls

:menusetup
title 部署－选择功能
cls
echo ---------------------------------------------------------------
echo     请选择任务。
echo ---------------------------------------------------------------
echo     [A]部署 Office 。
echo ---------------------------------------------------------------
echo     [Z] 返回上级。
echo     [Ctrl+C] 退出。
echo ---------------------------------------------------------------
echo.
:menusetupcho
set choice=
set /p choice=请输入对应的字母，按 [Enter] 开始:
if not "%Choice%"=="" SET Choice=%Choice:~0,1%
if /i "%choice%"=="a" goto menuoffice
if /i "%choice%"=="z" goto menu
echo 选择无效，请重新输入
echo.
goto menusetupcho
pause >nul
cls

:menuoffice
cls
title 部署 Office 365 / 2019－选择功能
echo ---------------------------------------------------------------
echo     请选择任务。
echo ---------------------------------------------------------------
echo     [A]在线安装 Office365 32位。
echo     [B]在线安装 Office365 64位。
echo     [C]在线安装 Office2019 32位。
echo     [D]在线安装 Office2019 64位。
echo     [E]使用官方安装工具在线安装 Office 365 。
echo     [F]下载 Office365 32位。
echo     [G]下载 Office365 64位。
echo     [H]下载 Office2019 32位。
echo     [I]下载 Office2019 64位。
echo     [J]关于 Office 更新通道说明。
echo     [K]使用官方工具移除 Office。
echo     [L]安装 Office 语言附件包 32位。
echo     [M]安装 Office 语言附件包 64位。
echo ---------------------------------------------------------------
echo     [Z] 返回上级。
echo     [Ctrl+C] 退出。
echo ---------------------------------------------------------------
echo.
:menuofficecho
set choice=
set /p choice=请输入对应的字母，按 [Enter] 开始:
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
if /i "%choice%"=="m" goto start13
if /i "%choice%"=="z" goto menusetup
echo 选择无效，请重新输入
echo.
goto menuofficecho
pause >nul
cls

:start1
cls
title 在线安装Office365 32位
cd /d .\files\
cmd /k "Setup Office365-x86.bat"
title 完成!
echo 完成!
echo. & pause
goto menuoffice

:start2
cls
title 在线安装Office365 64位
cd /d .\files\
cmd /k "Setup Office365-x64.bat"
title 完成!
echo 完成!
echo. & pause
goto menuoffice

:start3
cls
title 在线安装Office2019 32位
cd /d .\files\
cmd /k "Setup Office2019Enterprise-x86.bat"
title 完成!
echo 完成!
echo. & pause
goto menuoffice

:start4
cls
title 在线安装Office2019 64位
cd /d .\files\
cmd /k "Setup Office2019Enterprise-x64.bat"
title 完成!
echo 完成!
echo. & pause
goto menuoffice

:start5
cls
title 使用官方安装工具在线安装Office 365
echo 使用官方安装工具在线安装Office 365
echo.
echo 说明：工具为32位，打开后32位系统会直接安装，64位请根据提示自动安装64位
echo. & pause
echo 请在弹出窗口完成操作
cd /d .\files\
cmd /k "setupo365homepremretail.x86.zh-cn_.exe"
title 完成!
echo 完成!
echo. & pause
goto menuoffice

:start6
cls
title 下载Office365 32位
cd /d .\files\
cmd /k "Download Office365-x86.bat"
title 完成!
echo 完成!
echo. & pause
goto menuoffice

:start7
cls
title 下载Office365 64位
cd /d .\files\
cmd /k "Download Office365-x64.bat"
title 完成!
echo 完成!
echo. & pause
goto menuoffice

:start8
cls
title 下载Office2019 32位
cd /d .\files\
cmd /k "Download Office2019Enterprise-x86.bat"
title 完成!
echo 完成!
echo. & pause
goto menuoffice

:start9
cls
title 下载Office2019 64位
cd /d .\files\
cmd /k "Download Office2019Enterprise-x64.bat"
title 完成!
echo 完成!
echo. & pause
goto menuoffice

:start10
cls
title 关于Office更新通道
echo     关于Office更新通道：
echo     使用本脚本安装默认为每月通道
echo     如需更改，请进入files文件夹编辑对应版本xml文件中的更新通道
echo     文件中会找到 Channel="Monthly"
echo     其中 Monthly 可以修改为下面【】中的通道
echo ---------------------------------------------------------------
echo     下面列出更新通道名称：
echo     Office 2019 企业长期版【PerpetualVL2019】
echo     半年通道【Broad】
echo     半年通道（定向）【Targeted】
echo     每月通道【Monthly】
echo     每月通道（定向）【Insiders】
echo     测试通道（预览体验计划）【InsiderFast】
echo     开发通道（内部测试）【Dogfood】
echo ---------------------------------------------------------------
echo     按任意键回上级菜单，[Ctrl+C]关闭
echo ---------------------------------------------------------------
echo. & pause
goto menuoffice

:start11
cls
title 使用官方工具移除Office
echo 请在弹出窗口完成操作
cd /d .\files\clean\
cmd /k "o15-ctrremove.diagcab"
title 完成!
echo 完成!
echo. & pause
goto menuoffice

:start12
cls
title 安装Office语言附件包 32位
echo 安装Office语言附件包 32位
echo. & pause
echo 请等待弹出窗口完成部署
cd /d .\files\
cmd /k "setuplanguagepack.x86.zh-cn_.exe"
title 完成!
echo 完成!
echo. & pause
goto menuoffice

:start13
cls
title 安装Office语言附件包 64位
echo 安装Office语言附件包 64位
echo. & pause
echo 请等待弹出窗口完成部署
cd /d .\files\
cmd /k "setuplanguagepack.x64.zh-cn_.exe"
title 完成!
echo 完成!
echo. & pause
goto menuoffice

:information
cls
title 信息
echo ---------------------------------------------------------------
echo     MSO_Setup_Helper
echo     更新日期：2021/6/23
echo     版本：v1.0.1
echo     (c) zhxy-CN, Released under the MIT License.
echo ---------------------------------------------------------------
echo     [Z] 返回上级。
echo     [Ctrl+C]退出。
echo ---------------------------------------------------------------
echo.
:informationcho
set choice=
set /p choice=请输入对应的字母，按 [Enter] 开始:
if not "%Choice%"=="" SET Choice=%Choice:~0,1%
if /i "%choice%"=="z" goto menu
echo 选择无效，请重新输入
echo.
goto informationcho
pause >nul
cls
