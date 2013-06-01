@echo off
set path=%path%;C:/Windows/Microsoft.NET/Framework/v4.0.30319;
set EnableNuGetPackageRestore=true

echo Building project...
msbuild src/ExcelStream.sln /nologo /v:q /p:Configuration=Release /t:Clean
msbuild src/ExcelStream.sln /nologo /v:q /p:Configuration=Release /clp:ErrorsOnly

echo Merging assemblies...
if exist "publish" rmdir /s /q "publish"
mkdir publish\bin
mkdir publish\pkg
bin\ILMerge.exe /keyfile:src\ExcelStream.snk /internalize /wildcards /target:library ^
 /targetplatform:"v4,C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0" ^
 /out:"publish\bin\ExcelStream.dll" ^
 "src/ExcelStream/bin/Release/ExcelStream.dll" ^
 "src/ExcelStream/bin/Release/Ionic.Zip.Reduced.dll"
 
echo Creating NuGet packages...
for /r %%i in (src\packages\ExcelStream*.nuspec) do src\.nuget\nuget.exe pack %%i -symbols -OutputDirectory publish\pkg

echo Done.