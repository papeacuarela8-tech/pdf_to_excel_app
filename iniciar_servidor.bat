@echo off
title Servidor Papelería Acuarela - PDF a Excel
echo ================================================================
echo        INICIANDO CONVERTIDOR DE FACTURAS - PAPELERIA ACUARELA
echo ================================================================
echo.
echo [1/2] Abriendo el navegador en http://localhost:5000...
start http://localhost:5000
echo.
echo [2/2] Iniciando el servidor local de Python (Flask)...
echo (Por favor, deja esta ventana abierta mientras uses la aplicacion)
echo.
python app.py
pause
