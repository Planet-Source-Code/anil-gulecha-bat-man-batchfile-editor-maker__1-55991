@echo off
:MAIN
Cls
Echo �� Options ������������������������������������������������������������������ͻ
Echo �                                                                             �
Echo �                                 1.View Date                                 �
Echo �                                 2.View time                                 �
Echo �                                 3.Exit                                      �
Echo �                                                                             �
Echo �����������������������������������������������������������������������������ͼ
Choice /c:123 Choose an option
If ERRORLEVEL 3 goto end
If ERRORLEVEL 2 goto tim
If ERRORLEVEL 1 goto dat

:DAT
Echo �������������������������������������������������������������������������������
echo. | date | find "Cu"
Echo �������������������������������������������������������������������������������
Echo.
Pause
goto main
Exit

:TIM
Echo �������������������������������������������������������������������������������
echo. | time | find "Cu"
Echo �������������������������������������������������������������������������������
Echo.
Pause
goto main
Exit

:END
