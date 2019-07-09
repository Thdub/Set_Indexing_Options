# Set Indexing Options
Set indexed locations under "Indexing Options" (setting in control panel)

Group policy allows to include or exclude locations to indexed paths, but it won't override defaults or already set indexing options. 

3 different scripts:
  - "Set_Indexing_To_ StartMenus_Only.bat" and "Set_Indexing_To_ WinStartMenu_Only.bat" are more for post-install/deployment.
  1st one will add both Start Menu locations, 2nd one only "%ProgramData%\Microsoft\Windows\Start Menu". 
  
  - Set_Indexing_Options.bat is an interactive script (with a fancy browser) which allows to set your desired location(s) for search indexing.


Note : Needs Microsoft.Search.Interop.dll next to the script.

Download Microsoft.Search.Interop.dll : https://anonfile.com/FfA0Nfndn6/Microsoft.Search.Interop_dll

Download executable : https://github.com/Thdub/Set_Indexing_Options/releases/download/v1.0/Set_Indexing_Options.exe
