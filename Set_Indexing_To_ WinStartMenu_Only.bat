@echo off

setlocal
set start=%time%
set "scriptname=%TEMP%\SearchScopeTask.ps1"
set "lock=%TEMP%\wait%random%.lock"
cd /d "%~dp0"

:: Get SID
	for /f "tokens=1,2 delims==" %%s IN ('wmic path win32_useraccount where name^='%username%' get sid /value ^| find /i "SID"') do set "SID=%%t"

:: Make PS Script
	:: Call catalog
	@echo Add-Type -path "%~dp0Microsoft.Search.Interop.dll">>%scriptname%
	@echo $sm = New-Object Microsoft.Search.Interop.CSearchManagerClass>>%scriptname%
	@echo $catalog = $sm.GetCatalog^("SystemIndex"^)>>%scriptname%
	@echo $crawlman = $catalog.GetCrawlScopeManager^(^)>>%scriptname%
	:: Reset rules
	@echo $crawlman.RevertToDefaultScopes^(^)>>%scriptname%
	@echo $crawlman.SaveAll^(^)>>%scriptname%
	:: Remove rules
	@echo $crawlman.RemoveDefaultScopeRule^("file:///C:\Users\*"^)>>%scriptname%
	@echo $crawlman.SaveAll^(^)>>%scriptname%
	@echo $crawlman.RemoveDefaultScopeRule^("file:///C:\ProgramData\Microsoft\Windows\Start Menu\*"^)>>%scriptname%
	@echo $crawlman.SaveAll^(^)>>%scriptname%
	@echo $crawlman.RemoveDefaultScopeRule^("file:///C:\Users\*\AppData\Local\Microsoft\Windows\Temporary Internet Files\*"^)>>%scriptname%
	@echo $crawlman.SaveAll^(^)>>%scriptname%
	@echo $crawlman.RemoveDefaultScopeRule^("file:///C:\Users\*\AppData\Local\Temp\*"^)>>%scriptname%
	@echo $crawlman.SaveAll^(^)>>%scriptname%
	@echo $crawlman.RemoveDefaultScopeRule^("file:///C:\Users\*\AppData\*"^)>>%scriptname%
	@echo $crawlman.SaveAll^(^)>>%scriptname%
	@echo $crawlman.RemoveDefaultScopeRule^("iehistory://{%SID%}"^)>>%scriptname%
	@echo $crawlman.SaveAll^(^)>>%scriptname%
	:: Add start menu path
	@echo $crawlman.AddUserScopeRule^("file:///%ProgramData%\Microsoft\Windows\Start Menu\*",$true,$false,$null^)>>%scriptname%
	@echo $crawlman.SaveAll^(^)>>%scriptname%
	:: Remove automatically added favorites
	@echo $crawlman.RemoveDefaultScopeRule^("file:///%UserProfile%\Favorites\*"^)>>%scriptname%
	@echo $crawlman.SaveAll^(^)>>%scriptname%
	:: Reindex	
	@echo $Catalog.Reindex^(^)>>%scriptname%
	:: Delete lock file
	@echo Remove-Item "%lock%">>%scriptname%

:: Execute Task
	:: Create lock file
	@echo Locked>"%lock%"
	:: Launch ps script with admin and hidden
	PowerShell -NoProfile -ExecutionPolicy Bypass -c "& {Start-Process Powershell -ArgumentList '-ExecutionPolicy Unrestricted -File "%scriptname%" -force' -Verb RunAs -WindowStyle hidden}" >NUL 2>&1
	:: loop until lock file is deleted	
	:Wait
	if exist "%lock%" goto :Wait

:: Clean
	del "%scriptname%" /s /q >NUL 2>&1
exit /b