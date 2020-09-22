@echo off
:MAIN
Cls
Echo ÚÄ Formatter ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
Echo ³                                                                             ³
Echo ³                    BAT-Man Formatter Made using BAT-Man                     ³
Echo ³                                   1. A:                                     ³
Echo ³                                   2. B:                                     ³
Echo ³                                   3.Exit                                    ³
Echo ³                                                                             ³
Echo ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
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
Echo Ûß Error ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßÛ
Echo Û                                                                             Û
Echo Û                 An error occured while formatting the disc                  Û
Echo Û                   Press a key to return to the main menu                    Û
Echo Û                                                                             Û
Echo ÛÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÛ
Pause > nul
goto main

:END
cls
echo Thank you for using me.
