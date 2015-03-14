# file-size
This is a simple ATL shell extension which will subclass the status bar on each explorer file view window in order to display the size of the selected file(s). Version 0.5.2 cleans up the methods used to hook into the explorer view and adds two shortcut keys which I find useful: Escape (for deselect-all) and Ctrl+U (for navigate up one level). 

There are project files for Visual Studio 2013, a simple InnoSetup installer script and a binary installer supporting both 32-bit and 64-bit operating systems. 

If you cannot build this from source and wish to download the binary installer, please look in the master/install directory, or click here: https://github.com/christopherjr/file-size/blob/master/install/SetupStatusBarExtension_0.5.2.exe?raw=true 

The MD5 sum for the linked installer is 83aa66ee2c2da2a448d53e2eb654c2eb. 

Also note that you must show the Status Bar in order for this extension to show you the file size information. The Status Bar can be shown by opening an explorer window then checking the "Status Bar" item on the View menu. 
