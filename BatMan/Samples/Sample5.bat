@echo off
:MAIN
Cls
Echo �� Formatter ����������������������������������������������������������������Ŀ
Echo �                                                                             �
Echo �                    BAT-Man Formatter Made using BAT-Man                     �
Echo �                                   1. A:                                     �
Echo �                                   2. B:                                     �
Echo �                                   3.Exit                                    �
Echo �                                                                             �
Echo �������������������������������������������������������������������������������
echo.
Choice /c:abe Choose an option
If ERRORLEVEL 3 goto end
If ERRORLEVEL 2 goto bb
If ERRORLEVEL 1 goto aa
Exit

:AA
cls
format A: /u /v:EBD /autotest
if not errorlevel 0 goto F_ERR
Press a key to goto main menu
pause >nul
goto main

:BB
cls
format B: /u /v:EBD /autotest
if not errorlevel 0 goto F_ERR
Press a key to goto main menu
Pause >nul
goto main

:F_ERR
Echo �� Error ����������������������������������������������������������������������
Echo �                                                                             �
Echo �                 An error occured while formatting the disc                  �
Echo �                   Press a key to return to the main menu                    �
Echo �                                                                             �
Echo �������������������������������������������������������������������������������
Pause > nul
goto main

:END
cls
echo Thank you for using me.
