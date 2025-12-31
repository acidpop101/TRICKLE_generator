@echo off
REM Run Trickle Feed Generator (Excel -> Text)
REM Usage: run_gen.bat

echo Clean old output...
del /Q "d:\TRICKLE_generator\trickle_feed\cbs_output\*.txt"

echo Running Generator...
python "d:\TRICKLE_generator\trickle_feed\main.py" "d:\TRICKLE_generator\atm_dispute_app\atm_data.xlsx" "d:\TRICKLE_generator\trickle_feed\cbs_output"

echo.
echo Check the output folder: "d:\TRICKLE_generator\trickle_feed\cbs_output"
pause
