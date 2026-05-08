@echo off
setlocal

set WPS_PATH=C:\Program Files (x86)\Kingsoft\WPS Office\12.8.2.19315\office6

echo 正在注册WPS COM组件...

if exist "%WPS_PATH%\wps.exe" (
    regsvr32 "%WPS_PATH%\wps.exe" /s
    if %errorlevel% equ 0 (
        echo WPS COM组件注册成功
    ) else (
        echo WPS COM组件注册失败
    )
) else (
    echo 未找到WPS安装路径
    echo 请手动查找WPS安装目录并修改此脚本
)

if exist "%WPS_PATH%\et.exe" (
    regsvr32 "%WPS_PATH%\et.exe" /s
    if %errorlevel% equ 0 (
        echo ET COM组件注册成功
    ) else (
        echo ET COM组件注册失败
    )
)

pause