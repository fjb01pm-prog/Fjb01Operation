@echo off
setlocal enabledelayedexpansion
title FJB Operations Control - Auto Deploy System
color 0B

echo ============================================================
echo   FJB OPERATIONS CONTROL SYSTEM - AUTO DEPLOY AUTOMATION
echo ============================================================
echo.

echo [1/3] Sinkronisasi file index.html dan index.txt...
copy /Y "index.html" "index.txt" >nul
if %errorlevel% neq 0 (
    echo [ERROR] Gagal menyalin index.html ke index.txt
    goto error
)
echo [OK] File lokal berhasil disinkronkan.
echo.

echo [2/3] Mengirim script terbaru ke Google Apps Script...
echo       Script ID: 1_bQAhI52jUE0sahZ4mVeXS3dihMoffL8LmQPJY5KbGKV9lW0jJvP9_iM
call npx @google/clasp push --force
if %errorlevel% neq 0 (
    echo [WARNING] Clasp push mengalami kendala atau membutuhkan verifikasi login.
) else (
    echo [OK] Script berhasil dikirim ke Google Apps Script!
)
echo.

echo [3/3] Menyimpan ke Git dan memicu Auto-Deploy ke Vercel...
git add .
git diff --cached --quiet
if %errorlevel% neq 0 (
    git commit -m "chore(deploy): auto-deploy update %date% %time%"
) else (
    echo Tidak ada perubahan file baru untuk di-commit.
)

echo Mengunggah ke GitHub (Vercel Auto-Deployment)...
git push origin main
if %errorlevel% neq 0 (
    echo [ERROR] Gagal melakukan git push ke GitHub.
    goto error
)
echo [OK] Git push berhasil. Vercel sedang melakukan deploy otomatis!
echo.

echo ============================================================
echo   SEMUA TAHAPAN AUTO DEPLOY SELESAI DENGAN SUKSES!
echo   WebApp Live: https://fjb01-operation-001.vercel.app
echo ============================================================
echo.
pause
exit /b 0

:error
echo.
echo ============================================================
echo   [!] TERJADI KESALAHAN SAAT PROSES AUTO DEPLOY
echo ============================================================
pause
exit /b 1
