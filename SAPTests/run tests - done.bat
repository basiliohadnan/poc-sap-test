@echo off
echo 1. Executar todos os testes.
echo 2. Retomar exec.
echo.
set /p executionType=Defina a exec. desejada (1 ou 2):

if "%executionType%"=="1" (
    echo Running all tests...
    dotnet test --filter "TestCategory=done" "C:\Users\basil\OneDrive\Documentos\GitHub\LIT\poc-sap-test\SAPTests\SAPTests.csproj" -p:NoWarn=true -p:WarningLevel=0 --logger trx
) else if "%executionType%"=="2" (
    echo Resuming execution...
    dotnet test --filter "TestCategory=done" "C:\Users\basil\OneDrive\Documentos\GitHub\LIT\poc-sap-test\SAPTests\SAPTests.csproj" -p:NoWarn=true -p:WarningLevel=0 --logger trx
) else (
    echo Invalid executionType. Exiting...
    exit /b 1
)
