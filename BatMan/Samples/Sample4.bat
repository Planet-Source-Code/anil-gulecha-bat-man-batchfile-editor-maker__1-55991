@echo off
:MAIN
Cls
Echo 浜 Options 様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様融
Echo �                                                                             �
Echo �                                 1.View Date                                 �
Echo �                                 2.View time                                 �
Echo �                                 3.Exit                                      �
Echo �                                                                             �
Echo 藩様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様様�
Choice /c:123 Choose an option
If ERRORLEVEL 3 goto end
If ERRORLEVEL 2 goto tim
If ERRORLEVEL 1 goto dat

:DAT
Echo 陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳�
echo. | date | find "Cu"
Echo 陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳�
Echo.
Pause
goto main
Exit

:TIM
Echo 陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳�
echo. | time | find "Cu"
Echo 陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳陳�
Echo.
Pause
goto main
Exit

:END
