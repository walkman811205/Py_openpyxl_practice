@echo off
:mm

echo ---
echo (1)製作表格 (2)抓取現股 (3)抓取權證 (4)抓取標的物 (5)抓取Top20_購 (6)抓取Top20_售
set /P number=Please input your option :
cd C:\Users\admin\Desktop\Bigthree
set "t=%time%"

if "%number%"=="1" (
call python create_execel.py
)^
else if "%number%"=="2" (
echo 資料較多請稍後...
call python get_stock.py
echo done
)^
else if "%number%"=="3" (
echo 請稍後...
call python get_warrant.py
echo done
)^
else if "%number%"=="4" (
echo 資料較多請稍後...
call python get_sm.py
echo done
)^
else if "%number%"=="5" (
echo 請稍後...
call python sm_top20_buy.py
echo done
)^
else if "%number%"=="6" (
echo 請稍後...
call python sm_top20_sale.py
echo done
)^

else echo ##### 請輸入(1)~(6)數字 #####

set "t1=%time%"
if "%t1:~,2%" lss "%t:~,2%" set "add=+24"
set /a "times=(%t1:~,2%-%t:~,2%%add%)*360000+(1%t1:~3,2%%%100-1%t:~3,2%%%100)*6000+(1%t1:~6,2%%%100-1%t:~6,2%%%100)*100+(1%t1:~-2%%%100-1%t:~-2%%%100)" ,"ss=(times/100)%%60","mm=(times/6000)%%60","hh=times/360000","ms=times%%100"
echo ---
echo 花費時間 %hh%:%mm%:%ss%
echo ---
pause
goto mm



