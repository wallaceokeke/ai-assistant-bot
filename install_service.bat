@echo off
echo Installing KATE Assistant service...
python kate_service.py install
echo Starting KATE Assistant service...
python kate_service.py start
echo KATE Assistant service has been installed and started.
pause 