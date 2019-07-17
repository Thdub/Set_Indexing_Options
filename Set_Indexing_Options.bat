@if (@CodeSection == @Batch) @then

@echo off
setlocal disableDelayedExpansion

	call :conSize 100 30 100 9999

:: Title
	echo [97m]0;Indexing Options v2.0[97m[?25l
	cls
	echo                               =============================
	echo                               Indexing Options Script v2.1
	echo                               =============================
:: Written by Th. Dub @ResonantStep - 2019

	set "Clean=OFF"
	set "Idx_Tmp_Folder=%TEMP%\Indexing_Options_%random%.tmp"
	set "Idx_lock=%Idx_Tmp_Folder%\wait%random%.lock"
	set "Idx_scriptname=%Idx_Tmp_Folder%\SearchScopeTask.ps1"

:Indexing_Options
	set "Index=0"
	set "IndexedFolder="
	set "Increment_Index="
	set "line_up=2" & set "line_up2=2"
	set "Idx_1=-1" & set "Idx_2=-2" & set "Idx_3=-3" & set "Idx_4=-4" & set "Idx_5=-5" & set "Idx_6=-6" & set "Idx_7=-7" & set "Idx_8=-8" & set "Idx_9=-9" & set "Idx_10=-10"
	set "Idx_11=-11" & set "Idx_12=-12" & set "Idx_13=-13" & set "Idx_14=-14" & set "Idx_15=-15" & set "Idx_16=-16" & set "Idx_17=-17" & set "Idx_18=-18" & set "Idx_19=-19" & set "Idx_20=-20"
	set "IndexedFolder_1=" & set "IndexedFolder_2=" & set "IndexedFolder_3=" & set "IndexedFolder_4=" & set "IndexedFolder_5=" & set "IndexedFolder_6=" & set "IndexedFolder_7=" & set "IndexedFolder_8=" & set "IndexedFolder_9=" & set "IndexedFolder_10="
	set "IndexedFolder_11=" & set "IndexedFolder_12=" & set "IndexedFolder_13=" & set "IndexedFolder_14=" & set "IndexedFolder_15=" & set "IndexedFolder_16=" & set "IndexedFolder_17=" & set "IndexedFolder_18=" & set "IndexedFolder_19=" & set "IndexedFolder_20="

:: Menu
	echo:
	echo 1. Set custom locations& echo:
	echo 2. Add Windows start menus only& echo:
	echo 3. Remove all locations from indexing options& echo:
	echo 4. Default indexing options settings& echo: & echo:
	<nul set /p DummyName=Select your option, or 0 to exit: [?25h
	
	choice /c 12340 /n /m "" >nul 2>&1

	if errorlevel 5 ( set "Clean=Clean_OFF" & goto :Index_Task_Clean )

	if errorlevel 4 (
		echo [?25l4
		echo [14D[4A[92m4. Default indexing options settings[97m[4B
		set "Style=default"
		goto :Indexing_Options_Task
	)

	if errorlevel 3 (
		echo [?25l3
		echo [14D[6A[92m3. Remove all locations from indexing options[97m[6B
		set "Style=reset"
		goto :Indexing_Options_Task
	)

	if errorlevel 2 (
		echo [?25l2
		echo [14D[8A[92m2. Add Windows start menus only[97m[8B
		set "Style=startmenus"
		goto :Indexing_Options_Task
	)

	if errorlevel 1 (
		echo [?25l1
		echo [10A[92m1. Set custom locations[97m[10B
		set "Style=custom"
		goto :PathSelection
	)

:PathSelection
	setlocal EnableDelayedExpansion
	for /F "delims=" %%a in ('CScript //nologo //E:JScript "%~F0" "Select the folder or type the path you want to index, then click OK."') do (
	if %Index%==0 ( set "IndexedFolder=%%a" ) else ( set "IndexedFolder_%Index%=%%a" ))

	if "%IndexedFolder%" == "" (
		<nul set /p DummyName=[%line_up2%ASelect your option, or 0 to exit: [2X[10A
		endlocal && goto :Indexing_Options
	)

	if "%Index%" == "0" (
		echo You selected "%IndexedFolder%"
		set "Increment_Index=incr"
		goto :SelectMorePaths
	)

	if "!IndexedFolder_%Index%!" == "" (
		echo [%line_up%A[140X
		set "IndexedFolder="
		set "Filler=%line_up%"
		set /a "Filler-=1"

:Filler_Loop
		if "%Filler%" == "0" ( goto :Filler_Loop_End )
		echo [140X
		set /a "Filler-=1"
		goto :Filler_Loop

:Filler_Loop_End
		set /a "line_up2=%line_up2%+%line_up%"
		<nul set /p DummyName=[%line_up2%ASelect your option, or 0 to exit: [2X[10A
		endlocal && goto :Indexing_Options
	)

	if not "!IndexedFolder_%Index%!" == "%IndexedFolder%" (
	if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_1%!" ( if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_2%!" (
	if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_3%!" ( if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_4%!" (
	if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_5%!" ( if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_6%!" (
	if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_7%!" ( if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_8%!" (
	if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_9%!" ( if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_10%!" (
	if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_11%!" ( if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_12%!" (
	if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_13%!" ( if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_14%!" (
	if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_15%!" ( if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_16%!" (
	if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_17%!" ( if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_18%!" (
	if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_19%!" ( if not "!IndexedFolder_%Index%!" == "!IndexedFolder_%Idx_20%!" (

	echo You selected "!IndexedFolder_%Index%!"
	set /a "line_up+=2"
	set "Increment_Index=incr"
	goto :SelectMorePaths )))))))))))))))))))))

	set "IndexedFolder_%Index%="
	set "Increment_Index=no_incr"
	set /a "line_up+=1"
	goto :SelectMorePaths

:SelectMorePaths
	<nul set /p DummyName=Do you want to add another path to indexed locations? [Y/N][?25h
	choice /C:YN /M "" >nul 2>&1
	if errorlevel 2 ( echo [31mNo[97m& goto :PathResult )
	if "%Increment_Index%" == "incr" (
		set /a "Index+=1"
		set /a "Idx_1+=1" & set /a "Idx_2+=1" & set /a "Idx_3+=1" & set /a "Idx_4+=1" & set /a "Idx_5+=1" & set /a "Idx_6+=1" & set /a "Idx_7+=1"
		set /a "Idx_8+=1" & set /a "Idx_9+=1" & set /a "Idx_10+=1" & set /a "Idx_11+=1" & set /a "Idx_12+=1" & set /a "Idx_13+=1" & set /a "Idx_14+=1"
		set /a "Idx_15+=1" & set /a "Idx_16+=1" & set /a "Idx_17+=1" & set /a "Idx_18+=1" & set /a "Idx_19+=1" & set /a "Idx_20+=1"
	)
	echo [92mYes[97m[?25l
	goto :PathSelection

:PathResult
	echo:
	if "%Index%" == "0" (
		echo Indexed location is "%IndexedFolder%"
		set "More_Paths=Skip"
		goto :Indexing_Options_Task
	)
	if "%Index%" == "1" ( if "!IndexedFolder_%Index%!" == "" (
		echo Indexed location is "%IndexedFolder%"
		set "More_Paths=Skip"
		goto :Indexing_Options_Task
	))
	echo Indexed locations are
	echo "%IndexedFolder%"
	set /a "Count=%Index%"
	if %Index% GTR 0 ( if "!IndexedFolder_%Index%!" == "" ( set /a "Count-=1" ))

:ResultLoop
	if "%Count%" == "0" ( goto :Indexing_Options_Task )
	set "Index2=!IndexedFolder_%Count%!"
	echo "%Index2%"
	set /a "Count-=1"
	goto :ResultLoop

:Indexing_Options_Task
	if "%Style%" == "custom" ( echo: )
	mkdir "%Idx_Tmp_Folder%" >nul 2>&1
	set "Clean=Clean_ON"
	<nul set /p DummyName=Setting indexing options...[?25h

:: Get SID
	for /f "tokens=1,2 delims==" %%s IN ('wmic path win32_useraccount where name^='%username%' get sid /value ^| find /i "SID"') do set "SID=%%t"

:: Make PS Script
	@echo $host.ui.RawUI.WindowTitle = "Indexing Options v2.0 | Powershell Script">>%Idx_scriptname%
	@echo Add-Type -path "%~dp0Microsoft.Search.Interop.dll">>%Idx_scriptname%
	@echo $sm = New-Object Microsoft.Search.Interop.CSearchManagerClass>>%Idx_scriptname%
	@echo $catalog = $sm.GetCatalog^("SystemIndex"^)>>%Idx_scriptname%
	@echo $crawlman = $catalog.GetCrawlScopeManager^(^)>>%Idx_scriptname%
	@echo $crawlman.RevertToDefaultScopes^(^)>>%Idx_scriptname%
	@echo $crawlman.SaveAll^(^)>>%Idx_scriptname%
	if "%Style%" == "default" ( goto :MakeDefault )
	@echo $crawlman.RemoveDefaultScopeRule^("file:///C:\Users\*"^)>>%Idx_scriptname%
	@echo $crawlman.SaveAll^(^)>>%Idx_scriptname%
	@echo $crawlman.RemoveDefaultScopeRule^("file:///C:\ProgramData\Microsoft\Windows\Start Menu\*"^)>>%Idx_scriptname%
	@echo $crawlman.SaveAll^(^)>>%Idx_scriptname%
	@echo $crawlman.RemoveDefaultScopeRule^("file:///C:\Users\*\AppData\Local\Microsoft\Windows\Temporary Internet Files\*"^)>>%Idx_scriptname%
	@echo $crawlman.SaveAll^(^)>>%Idx_scriptname%
	@echo $crawlman.RemoveDefaultScopeRule^("file:///C:\Users\*\AppData\Local\Temp\*"^)>>%Idx_scriptname%
	@echo $crawlman.SaveAll^(^)>>%Idx_scriptname%
	@echo $crawlman.RemoveDefaultScopeRule^("file:///C:\Users\*\AppData\*"^)>>%Idx_scriptname%
	@echo $crawlman.SaveAll^(^)>>%Idx_scriptname%
	@echo $crawlman.RemoveDefaultScopeRule^("iehistory://{%SID%}"^)>>%Idx_scriptname%
	@echo $crawlman.SaveAll^(^)>>%Idx_scriptname%
	if "%Style%" == "default" ( goto :MakeDefault )
	if "%Style%" == "reset" ( goto :Finish_Ps )
	if "%Style%" == "startmenus" ( goto :AddStartMenus )
	if "%Style%" == "custom" ( goto :SetCustomPaths )

:MakeDefault
	@echo $crawlman.AddUserScopeRule^("file:///C:\Users\*",$true,$false,$null^)>>%Idx_scriptname%
	@echo $crawlman.SaveAll^(^)>>%Idx_scriptname%
	@echo $crawlman.AddUserScopeRule^("file:///C:\ProgramData\Microsoft\Windows\Start Menu\*",$true,$false,$null^)>>%Idx_scriptname%
	@echo $crawlman.SaveAll^(^)>>%Idx_scriptname%
	@echo $crawlman.AddUserScopeRule^("iehistory://{%SID%}",$true,$false,$null^)>>%Idx_scriptname%
	@echo $crawlman.SaveAll^(^)>>%Idx_scriptname%
	goto :Reindex

:AddStartMenus
	@echo $crawlman.AddUserScopeRule^("file:///%ProgramData%\Microsoft\Windows\Start Menu\Programs\*",$true,$false,$null^)>>%Idx_scriptname%
	@echo $crawlman.SaveAll^(^)>>%Idx_scriptname%
	@echo $crawlman.AddUserScopeRule^("file:///%AppData%\Microsoft\Windows\Start Menu\Programs\*",$true,$false,$null^)>>%Idx_scriptname%
	@echo $crawlman.SaveAll^(^)>>%Idx_scriptname%
	goto :Finish_Ps

:SetCustomPaths
	@echo $crawlman.AddUserScopeRule^("file:///%IndexedFolder%\*",$true,$false,$null^)>>%Idx_scriptname%
	@echo $crawlman.SaveAll^(^)>>%Idx_scriptname%
	if "%More_Paths%" == "Skip" ( goto :Finish_Ps )

:MorePathsLoop
	if "%Index%" == "0" ( goto :Finish_Ps )
	if %Index% GTR 0 ( if "!IndexedFolder_%Index%!" == "" ( set /a "Index-=1" ))
	set "Index2=!IndexedFolder_%Index%!"
	@echo $crawlman.AddUserScopeRule^("file:///%Index2%\*",$true,$false,$null^)>>%Idx_scriptname%
	@echo $crawlman.SaveAll^(^)>>%Idx_scriptname%
	set /a "Index-=1"
	goto :MorePathsLoop

:Finish_Ps
	@echo $crawlman.RemoveDefaultScopeRule^("file:///%UserProfile%\Favorites\*"^)>>%Idx_scriptname%
	@echo $crawlman.SaveAll^(^)>>%Idx_scriptname%

:Reindex
	@echo $Catalog.Reindex^(^)>>%Idx_scriptname%

:: Backup script before lock	
	copy /b /v /y "%Idx_scriptname%" "%TEMP%\SearchScopeTask2.ps1" >nul 2>&1
	@echo Remove-Item "%Idx_lock%">>%Idx_scriptname%
:: Execute Task
	@echo Locked>"%Idx_lock%"
	PowerShell -NoProfile -ExecutionPolicy Unrestricted -File "%Idx_scriptname%" -force >nul 2>&1

:Wait
	if exist "%Idx_lock%" ( goto :Wait )

:Index_Task_Clean
	if "%Clean%" == "Clean_OFF" (
		echo:
		echo [?25l[93mNo indexing location has been set.[97m
		goto :End_Pause
	)
	echo [?25l[92mDone.[97m
	echo [93mIndexing options setting task has completed successfully.[97m

:: Backup Script
	move /y	"%TEMP%\SearchScopeTask2.ps1" "%~dp0SearchScope_Task_Backup.ps1" >nul 2>&1

:CleanMore_1
	del /F /Q /S "%Idx_scriptname%" >nul 2>&1
	if not exist "%Idx_scriptname%" ( goto :CleanMore_2 ) else ( goto :CleanMore_1 )

:CleanMore_2
	rmdir "%Idx_Tmp_Folder%\" /s /q >nul 2>&1
	if not exist "%Idx_Tmp_Folder%\" ( goto :End_Pause ) else ( goto :CleanMore_2 )
	
:End_Pause
	timeout /t 3 /nobreak >NUL 2>&1
exit /b

::============================================================================================================
:conSize  winWidth  winHeight  bufWidth  bufHeight
::============================================================================================================
	mode con: cols=%1 lines=%2
	Powershell "&{$H=get-host;$W=$H.ui.rawui;$B=$W.buffersize;$B.width=%3;$B.height=%4;$W.buffersize=$B;}"
	goto :eof
::============================================================================================================

@end


var shl = new ActiveXObject("Shell.Application");
var folder = shl.BrowseForFolder(0, WScript.Arguments(0), 0x00000050,17);
WScript.Stdout.WriteLine(folder ? folder.self.path : "");