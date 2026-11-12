@echo off
echo ========================================================
echo Iniciando Google Chrome em Modo Homologado (Porta 9222)
echo ========================================================
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="%LOCALAPPDATA%\Google\Chrome\User Data" https://conectividadesocialv2.caixa.gov.br/sicns/
echo Chrome iniciado!
