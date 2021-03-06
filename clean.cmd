@echo off
setlocal
:: Note: We've disabled node reuse because it causes file locking issues.
:: The issue is that we extend the build with our own targets which
:: means that that rebuilding cannot successully delete the task
:: assembly.
:: Check prerequisites

set _msbuildexe="%ProgramFiles(x86)%\MSBuild\12.0\Bin\MSBuild.exe"
if not exist %_msbuildexe% set _msbuildexe="%ProgramFiles%\MSBuild\12.0\Bin\MSBuild.exe"
if not exist %_msbuildexe% set _msbuildexe="%ProgramFiles(x86)%\MSBuild\14.0\Bin\MSBuild.exe"
if not exist %_msbuildexe% set _msbuildexe="%ProgramFiles%\MSBuild\14.0\Bin\MSBuild.exe"
if not exist %_msbuildexe% echo Error: Could not find MSBuild.exe. Please see https://github.com/dotnet/corefx/blob/master/docs/Developers.md for build instructions. && goto :eof
:: Log build command line
set _buildprefix=echo
set _buildpostfix=^> "%~dp0msbuild.log"
call :build %*
:: Build
set _buildprefix=
set _buildpostfix=
call :build %*
goto :eof
:build
%_buildprefix% %_msbuildexe% "%~dp0CalendarSync.sln" /t:Clean /p:Configuration=Release
goto :eof
