set signtool=C:\Work\BuildTools\Certificates\codesign.bat /du "https://www.unicontsoft.com"
%signtool% /d "Unicontsoft SQL Monitor" %~dp0UcsSqlMonitor.exe
