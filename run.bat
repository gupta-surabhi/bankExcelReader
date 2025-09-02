@echo off
REM Compile
javac -proc:none -cp "lib/*" -d lib src\BankExcelSegregator.java
echo Compiled successfully!!
echo.

echo Starting the program...
echo.
REM Run
java -cp "lib;lib/*" BankExcelSegregator
pause
