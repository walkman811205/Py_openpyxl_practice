@echo off
:mm

echo ---
echo (1)�s�@��� (2)����{�� (3)����v�� (4)����Ъ��� (5)���Top20_�� (6)���Top20_��
set /P number=Please input your option :
cd C:\Users\admin\Desktop\Bigthree
set "t=%time%"

if "%number%"=="1" (
call python create_execel.py
)^
else if "%number%"=="2" (
echo ��Ƹ��h�еy��...
call python get_stock.py
echo done
)^
else if "%number%"=="3" (
echo �еy��...
call python get_warrant.py
echo done
)^
else if "%number%"=="4" (
echo ��Ƹ��h�еy��...
call python get_sm.py
echo done
)^
else if "%number%"=="5" (
echo �еy��...
call python sm_top20_buy.py
echo done
)^
else if "%number%"=="6" (
echo �еy��...
call python sm_top20_sale.py
echo done
)^

else echo ##### �п�J(1)~(6)�Ʀr #####

set "t1=%time%"
if "%t1:~,2%" lss "%t:~,2%" set "add=+24"
set /a "times=(%t1:~,2%-%t:~,2%%add%)*360000+(1%t1:~3,2%%%100-1%t:~3,2%%%100)*6000+(1%t1:~6,2%%%100-1%t:~6,2%%%100)*100+(1%t1:~-2%%%100-1%t:~-2%%%100)" ,"ss=(times/100)%%60","mm=(times/6000)%%60","hh=times/360000","ms=times%%100"
echo ---
echo ��O�ɶ� %hh%:%mm%:%ss%
echo ---
pause
goto mm



