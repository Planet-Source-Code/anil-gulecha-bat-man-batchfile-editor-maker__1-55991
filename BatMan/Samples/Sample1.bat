@echo off
:main
Cls
Echo �� MS-DOS Helper created with BAT-Man batchfile maker �������������������������
Echo �                                                                             �
Echo �                    Choose the command you need help for:                    �
Echo �                                                                             �
Echo �                                 1.DIR                                       �
Echo �                                 2.COPY                                      �
Echo �                                 3.MOVE                                      �
Echo �                                 4.ATTRIB                                    �
Echo �                                 5.EXIT                                      �
Echo �                                                                             �
Echo �������������������������������������������������������������������������������
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
Echo �����������������������������������������������������������������������������ͻ
Echo �                                                                             �
Echo �                                  Thank you                                  �
Echo �                     This file was made in 10 mins using                     �
Echo �                           BAT-Man Batchfile maker                           �
Echo �                            (c) 2004 Anil Gulecha                            �
Echo �                               a.k.a GeekFreek                               �
Echo �                                                                             �
Echo �����������������������������������������������������������������������������ͼ
Exit
