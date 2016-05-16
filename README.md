# Modify_FolderIcon
When there's no file inside a folder, user powershell to set the Icon of folder to Grey
Otherwise, just keep the original folder Icon

User Guide：
1. run.bat,Monitor.ps1,folder.ico
2.  copy these 3 file to the top path of your folder
3.double click run.bat
4.then it'll keep monitor your folder



Weakness
1. sometimes the icon will be delay 1-15min to become grey
2. only detect unhidden files

to solve this Weakness 1:
everytime restart explorer , but another weakness comes:
the window will flash, and the grey icon will there immediately.
sometimes the first open window will not close automatically.


other:
if you want to change to another icon：
1. a image
2. download & install ICofx http://icofx.ro/downloads.html 
3. drap the image into icofx，-->edit-->，save as  folder.icon(the file name must be this one, or you could change code in Monitor.ps1 to your correspond file name)
4.your icon to replace folder.icon