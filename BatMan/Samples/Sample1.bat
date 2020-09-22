@echo off
:main
Cls
Echo Û฿ MS-DOS Helper created with BAT-Man batchfile maker ฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿Û
Echo Û                                                                             Û
Echo Û                    Choose the command you need help for:                    Û
Echo Û                                                                             Û
Echo Û                                 1.DIR                                       Û
Echo Û                                 2.COPY                                      Û
Echo Û                                 3.MOVE                                      Û
Echo Û                                 4.ATTRIB                                    Û
Echo Û                                 5.EXIT                                      Û
Echo Û                                                                             Û
Echo ÛÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÛ
echo.
echo Press a key
Choice /c:12345 Choose an option
If ERRORLEVEL 5 goto END
If ERRORLEVEL 4 goto attrib
If ERRORLEVEL 3 goto move
If ERRORLEVEL 2 goto copy
If ERRORLEVEL 1 goto dir
exit

:MOVE
cls
move/?
echo Hit a key to goto Main Menu
pause >nul
goto main

:COPY
cls
copy/?
echo Hit a key to goto Main Menu
pause >nul
goto main

:DIR
cls
dir/?
echo Hit a key to goto Main Menu
pause >nul
goto main

:ATTRIB
cls
attrib/?
echo Hit a key to goto Main Menu
pause >nul
goto main

:END
cls
echo.
echo.
echo.
echo.
Echo ษอออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออป
Echo บ                                                                             บ
Echo บ                                  Thank you                                  บ
Echo บ                     This file was made in 10 mins using                     บ
Echo บ                           BAT-Man Batchfile maker                           บ
Echo บ                            (c) 2004 Anil Gulecha                            บ
Echo บ                               a.k.a GeekFreek                               บ
Echo บ                                                                             บ
Echo ศอออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผ
Exit
