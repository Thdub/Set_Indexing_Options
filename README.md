# Set Indexing Options
- Script to set indexed locations under "Indexing Options"
- Group policy allows to include or exclude locations, but it won't override defaults or already set indexing options. 
- 3 different scripts:
  - "Set_Indexing_To_ StartMenus_Only" and "Set_Indexing_To_ WinStartMenu_Only" are more for post-install/deployment. 1st one will add both Start Menu locations, 2nd one only "%ProgramData%\Microsoft\Windows\Start Menu" as unique location, 
  - Set_Indexing_Options is an interactive script (with a fancy browser) which allows to set your desired location(s) for search
  indexing.
  
- Note: Needs Microsoft.Search.Interop.dll next to the script. Open the file with Powershell ISE.
- Download Microsoft.Search.Interop.dll : https://anonfile.com/FfA0Nfndn6/Microsoft.Search.Interop_dll
- Download Set_Indexing_Options.exe : 
