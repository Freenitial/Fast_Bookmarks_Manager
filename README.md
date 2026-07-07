# **Fast Bookmarks Manager**

**Graphical/Cmd tool for Chrome / Edge / Firefox / OperaGX bookmarks-bar export and import**

--------------------

### Features ✨ 

- 🔄 Export & import bookmarks-**bar** between browsers 
- 🌐 Support for Chrome / Edge / Firefox / OperaGX profiles
- 📂 Preserve folders structure during export/import
- 🗃️ Organized backup as .url files/folders structure, or HTML file
- 🔍 Treeview for selective bookmark management
- 🚀 Auto-restore functionality for scripting
- 📊 Profile-specific bookmark counting
- 📝 Logs

--------------------

### Tab 1: Backup 🗃️
- Select Chrome / Edge / Firefox / OperaGX profiles of bookmarks bar to export
- Tree view with checkboxes for selective backup
- Export url files, preserving folders structure
- Ignores browser-specific URLs (chrome://, edge://)

<img width="584" height="542" alt="Capture d&#39;écran 2026-07-07 185509" src="https://github.com/user-attachments/assets/fd5746c2-6c9b-42f3-a306-7a301a804c17" />

--------------------

### Tab 2: Restore 🌐
- Import .url files into browser bookmarks
- Profile-specific restoration capability
- Visual tree browser for selective import
- Auto-detect available profiles
- Import counter to track selected URLs
- Compatible with any valid .url file

<img width="584" height="542" alt="Capture d&#39;écran 2026-07-07 185853" src="https://github.com/user-attachments/assets/4293d64a-1f63-434c-973a-9e11297a2e29" />

--------------------

### Optional Arguments 🔧
   
1) `-autorestore`
   - Auto restore all urls founds from source folder or HTML file to each default profile of each browser

2) `-source "Full\Filepath\to\sourceFolder"`
   - Full path to the folder containing .url files for restoration. Default : Script folder
   - When using -source argument, not followed by HTML file or folder path, HTML file in script dir will be prioritized

3) `-logfile "Full\Filepath\to\logfile.log"`
   - Full path for the log file. Default : %temp%\Fast_Bookmarks_Manager.log


--------------------

### Script Usage 📝

_Requirement : Windows 10 build 1607 +_

To launch Fast Bookmarks Manager normally: **Just open .bat file**

To launch with auto-restore and specify source and logfile paths:  
```
start "" /d "SCRIPT_DIR" Fast_Bookmarks_Manager -autorestore "C:\Source_Dir" "C:\Logfile.log"
start "" /d "SCRIPT_DIR" Fast_Bookmarks_Manager -autorestore "C:\Source_File.html" "C:\Logfile.log"
```  

Multi-line example:
```
start "" /d "SCRIPT_DIR" Fast_Bookmarks_Manager ^
                         -autorestore ^
                         -source "C:\Path\To\Source" ^
                         -logfile "C:\Path\To\Logfile.log"
```  

--------------------
