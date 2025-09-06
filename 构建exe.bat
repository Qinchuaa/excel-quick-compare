@echo off
chcp 65001 >nul
echo Excel数据对比工具 - 可执行文件构建器
echo ================================================
echo.
echo 正在构建可执行文件，请稍候...
echo.

python build_exe.py

echo.
echo 构建完成！
echo.
pause