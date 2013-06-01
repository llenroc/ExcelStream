@echo off
call build.cmd
for /r %%i in (publish\pkg\ExcelStream*.nupkg) do src\.nuget\nuget.exe Push %%i