@echo off
rem ============================================================
rem  数据统计平台 一键发布到 GitHub Pages
rem  首次运行：输入仓库地址（须先在 github.com 建好空仓库 Public）
rem  以后每次改版：直接双击，自动提交并推送 index.html / mobile.html
rem  依赖：本机已安装 Git（git-scm.com）
rem ============================================================
cd /d "%~dp0"
setlocal enabledelayedexpansion

git --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo [错误] 本机没有安装 Git，或 Git 未添加到系统 PATH。
    echo 如果你刚刚才安装 Git：先关闭这个窗口，重新打开 cmd 再运行本脚本。
    echo.
    echo 请先下载安装：
    echo   https://git-scm.com/download/win
    echo.
    echo 安装时一路点 Next 即可；装完后重新打开 cmd 窗口，输入：
    echo   git --version
    echo 看到版本号表示安装成功。
    echo.
    echo 也可以不装 Git：打开《云端部署说明_GitHubPages.md》看第三节，
    echo 用 GitHub 网页直接上传 index.html + mobile.html。
    pause
    exit /b 1
)

set "URL_FILE=.git_remote_url.txt"
set "REMOTE="

if exist "%URL_FILE%" set /p REMOTE=<"%URL_FILE%"

if "%REMOTE%"=="" (
    echo.
    echo [首次运行] 请输入你的仓库地址，例如：
    echo   https://github.com/你的用户名/stats-platform.git
    echo 如果还没建仓库：先到 github.com 右上角 New repository
    echo 创建（Repository name 填 stats-platform，选 Public，
    echo 不要勾 Add a README file）。
    echo 直接回车则取消。
    echo.
    set /p REMOTE=仓库地址: 
    if "!REMOTE!"=="" exit /b
    echo !REMOTE!> "%URL_FILE%"
)

if not exist ".git" git init
git config user.email "deploy@local"
git config user.name "deploy"
git branch -M main >nul 2>&1
git remote remove origin >nul 2>&1
git remote add origin "%REMOTE%"
git add index.html mobile.html
git commit -m "deploy update"
git push -u origin main

if errorlevel 1 (
    echo.
    echo [推送失败] 常见原因与处理：
    echo   1 仓库非空或历史冲突 - 依次执行下面三条再重试：
    echo      git pull origin main --allow-unrelated-histories
    echo      git push -u origin main
    echo   2 首次推送会弹出 GitHub 登录窗口（Git Credential Manager），
    echo     若未弹出或登录失败，请再双击运行一次。
    pause
    exit /b 1
)

echo.
echo [完成] 已推送 index.html / mobile.html。
echo 等约 1 分钟，用手机浏览器打开线上地址刷新即可：
echo   https://你的用户名.github.io/stats-platform/mobile.html
pause
