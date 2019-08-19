set currentLocation=%~dp0
%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase "%currentLocation%\bin64\AccessReportGenerator.dll"
PAUSE