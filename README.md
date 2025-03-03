# **Fast Bookmarks Manager**

**Graphical PowerShell tool for browser bookmarks-bar export and import**

--------------------

### Features ✨ 

- 🔄 Export & import bookmarks-bar between browsers 
- 🌐 Full support for Chrome and Edge profiles
- 📂 Preserve folders structure during export/import
- 🗃️ Organized backup as .url files
- 🔍 Tree view for selective bookmark management
- ⚙️ Customizable target/source folders
- 🚀 Auto-restore functionality for scripting
- 📊 Profile-specific bookmark counting
- 📝 Logs

--------------------

### Tab 1: Backup 🗃️
- Select Chrome and Edge profiles of bookmarks bar to export
- Tree view with checkboxes for selective backup
- Export url files, preserving folders structure
- Ignores browser-specific URLs (chrome://, edge://)

![image](https://github.com/user-attachments/assets/e8ce5bc7-a53c-4e2b-afdc-0425a02fb8c9)

--------------------

### Tab 2: Restore 🌐
- Import .url files into browser bookmarks
- Profile-specific restoration capability
- Visual tree browser for selective import
- Auto-detect available profiles
- Import counter to track selected URLs
- Compatible with any valid .url file

![image](https://github.com/user-attachments/assets/76acd459-5d0b-4af0-943e-6ab7470991a3)

--------------------

### Tab 3: Settings ⚙️
- Configure backup target folder
- Set restore source folder

![image](https://github.com/user-attachments/assets/6ac6cafb-f249-4728-abe5-0f4ccd5d5224)

--------------------

### Optional Arguments 🔧
   
1) `-autorestore`
   - Enables automatic restoration mode without user prompts, for each default profile of each browser

2) `-source "Full\Filepath\to\sourceFolder"`
   - Full path to the folder containing .url files for restoration. Default : Script folder

3) `-logfile "Full\Filepath\to\logfile.log"`
   - Full path for the log file. Default : %temp%\Fast_Bookmarks_Manager.log


--------------------

### Script Usage 📝

_Requirement : Windows 10 build 1607 +_

To launch Fast Bookmarks Manager normally: just open .bat file  

To launch with auto-restore and specify source and logfile paths:  
```
start "" /d "SCRIPT_DIR" Fast_Bookmarks_Manager -autorestore "C:\Source_Dir" "C:\Logfile.log"
```  

Multi-line example:
```
start "" /d "SCRIPT_DIR" Fast_Bookmarks_Manager ^
                         -autorestore ^
                         -source "C:\Path\To\Source" ^
                         -logfile "C:\Path\To\Logfile.log"
```  

--------------------
