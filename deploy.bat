@echo off
echo ══════════════════════════════════════════
echo  Thomson EA Dashboard — Docker Deploy
echo ══════════════════════════════════════════

:: Switch Docker to Windows containers
echo [1/4] Switching to Windows containers...
"C:\Program Files\Docker\Docker\DockerCli.exe" -SwitchDaemon

:: Build image
echo [2/4] Building image...
docker build -t neuro-dashboard:latest .
if %errorlevel% neq 0 (
    echo ERROR: Build failed.
    pause
    exit /b 1
)

:: Stop old container if running
echo [3/4] Stopping old container (if any)...
docker stop neuro-dashboard 2>nul
docker rm   neuro-dashboard 2>nul

:: Run new container
echo [4/4] Starting container...
docker-compose up -d

echo.
echo ✅ Dashboard running at http://localhost:8501
echo    Access from phone: http://%COMPUTERNAME%:8501
echo    or use VPS public IP: http://YOUR_VPS_IP:8501
echo.
pause
