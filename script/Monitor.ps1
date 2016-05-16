################################################################################################
#Function: when add file/folder or delete file/folder from project folders, then need to       #
#check whether any files inside, if not, then set the folder icon as grey, otherwise, non-grey #
#Author:Rita Chen                                                                              #
#History                                                                              #
#-------------------------                                                                     #
#Change  Editor    Date                                                                        #
#initial RitaChen 2016-05-10                                                                   #
################################################################################################
#1. get current path of the script
$folder = Split-Path -Parent $MyInvocation.MyCommand.Definition
cd $folder

#2. passing param to event of watcher
$pso = new-object psobject -property @{
folder ="$folder"
IconfullPath = "$folder\folder.ico"
INIFullPath="$folder\DESKTOP.INI"
}

#3. set the system watcher to watcher the changes of the folder and subfolders
$FSWProps = @{
   Path = $folder
   IncludeSubdirectories = $true  
}
$FileSystemWatcher = New-Object System.IO.FileSystemWatcher -Property $FSWProps

#4. set the action function
$ChangeIconAction = {  
	
	##################################
	#functions for set and remove Icon
	##################################
	<#
.SYNOPSIS
This function sets a folder icon on specified folder.
.DESCRIPTION
This function sets a folder icon on specified folder. Needs the path to the icon file to be used and the path to the folder the icon is to be applied to. This function will create two files in the destination path, both set as Hidden files. DESKTOP.INI and FOLDER.ICO
.EXAMPLE
Set-FolderIcon -Icon "C:\Users\Mark\Downloads\Radvisual-Holographic-Folder.ico" -Path "C:\Users\Mark"
Changes the default folder icon to the custom one I donwloaded from Google Images.
.EXAMPLE
Set-FolderIcon -Icon "C:\Users\Mark\Downloads\wii_folder.ico" -Path "\\FAMILY\Media\Wii"
Changes the default folder icon to custom one for a UNC Path.
.EXAMPLE
Set-FolderIcon -Icon "C:\Users\Mark\Downloads\Radvisual-Holographic-Folder.ico" -Path "C:\Test" -Recurse
Changes the default folder icon to custom one for all folders in specified folder and that folder itself.
.NOTES 
Created by Mark Ince on May 4th, 2014. Contact me at mrince@outlook.com if you have any questions.
#>
function Set-FolderIcon
{
	[CmdletBinding()]
	param
	(	
		[Parameter(Mandatory=$True,
		Position=0)]
		[string[]]$Icon,
		[Parameter(Mandatory=$True,
		Position=1)]
		[string]$Path,
		[Parameter(Mandatory=$False)]
		[switch]
		$Recurse	
	)
	BEGIN
	{
		$originallocale = $PWD
		#Creating content of the DESKTOP.INI file.
		$ini = '[.ShellClassInfo]
				IconFile=FOLDER.ICO
				IconIndex=0
				ConfirmFileOp=0'
		Set-Location $Path
		Set-Location ..	
		Get-ChildItem | Where-Object {$_.FullName -eq "$Path"} | ForEach {$_.Attributes = 'Directory, System'}
	}	
	PROCESS
	{
		Remove-Item $Path\DESKTOP.INI -Force -errorAction SilentlyContinue -errorVariable errors
		$ini | Out-File $Path\DESKTOP.INI -Force -errorAction SilentlyContinue -errorVariable errors
		If ($Recurse -eq $True)
		{
			Copy-Item -Path $Icon -Destination $Path\FOLDER.ICO -Force -errorAction SilentlyContinue -errorVariable errors
			$recursepath = Get-ChildItem $Path -r | Where-Object {$_.Attributes -match "Directory"}
			ForEach ($folder in $recursepath)
			{
				Set-FolderIcon -Icon $Icon -Path $folder.FullName
			}
		
		}
		else
		{
			Copy-Item -Path $Icon -Destination $Path\FOLDER.ICO -Force -errorAction SilentlyContinue -errorVariable errors
		}	
	}	
	END
	{
		$inifile = Get-Item $Path\DESKTOP.INI
		#$inifile.Attributes = 'Hidden'
		attrib -A +S +H $inifile
		$icofile = Get-Item $Path\FOLDER.ICO
		$icofile.Attributes = 'Hidden'
		Set-Location $originallocale		
	}
}
<#

#>
function Remove-SetIcon
{
	[CmdletBinding()]
	param
	(	
		[Parameter(Mandatory=$True,
		Position=0)]
		[string]$Path
	)
	BEGIN
	{
		$originallocale = $PWD
		#$iconfiles = Get-ChildItem $Path -Recurse -Force | Where-Object {$_.Name -like "FOLDER.ICO"}
		#$iconfiles = $iconfiles.FullName
		#$inifiles = Get-ChildItem $Path -Recurse -Force | where-Object {$_.Name -like "DESKTOP.INI"}
		#$inifiles = $inifiles.FullName
		$iconfiles="$Path\FOLDER.ICO"
		$inifiles="$Path\DESKTOP.INI"
	}
	PROCESS
	{
		Remove-Item $iconfiles -Force -errorAction SilentlyContinue -errorVariable errors
		Remove-Item $inifiles -Force -errorAction SilentlyContinue -errorVariable errors
		Set-Location $Path
		Set-Location ..
		Get-ChildItem | Where-Object {$_.FullName -eq "$Path"} | ForEach {$_.Attributes = 'Directory'}	
	}
	END
	{
		Set-Location $originallocale
	}
}
	###########################
	#Start to set the grey Icon
	###########################
	$folder=$event.MessageData.folder
	$IconfullPath=$event.MessageData.IconfullPath
	$INIFullPath=$event.MessageData.INIFullPath
	$IconFile=Split-Path -Leaf $IconfullPath
	$INIFile=Split-Path -Leaf $INIFullPath
	$fullpath = $Event.SourceEventArgs.FullPath
	Write-Host "fullpath: $fullpath" -fore green
	$path=Split-Path -Parent $fullpath
	$fullname = $Event.SourceEventArgs.Name 
	$name=Split-Path -leaf  $fullname
	$changeType = $Event.SourceEventArgs.ChangeType 
	$timeStamp = $Event.TimeGenerated 
	Write-Host "folder:$folder,filename:$name,$changeType at $timeStamp" -fore green 
	Out-File -FilePath $folder\outlog.txt -Append -InputObject "folder:$folder,filename:$name,$changeType at $timeStamp"	
	#Write-Host "$name : $INIFile : $IconFile"
	if("$name" -eq "$INIFile" -or "$name" -eq "$IconFile"){
		#Write-Host "ini, icon"
	}else{
		#Write-Host "not ini, icon"
		#step 1: find all subfolders Recursely
		dir -r "$folder" | where { $_ -is [System.IO.DirectoryInfo] } | foreach-object  -process { 
		#Remove-SpecialFiles($_.FullName) 
		$subfolder=$_.FullName
		$directoryInfo=dir -r $subfolder | where { $_ -is [System.IO.FileInfo] } | Measure-Object
		$directoryInfo.count
		#step 2: check all subfolders, any files inside, if not, set Grey Icon, 
		#otherwise, check it's Grey or not, if Grey , remove it
		if($directoryInfo.count -eq 0)
		{		
			Write-Host "$subfolder is blank"
			#Out-File -FilePath $folder\log_all.txt -Append -InputObject "$subfolder is blank"
			#copy-item $folder\icon\DESKTOP.INI -destination $subfolder -Force -errorAction SilentlyContinue -errorVariable errors
			if (![System.IO.File]::Exists("$subfolder\$INIFile")){
				Set-FolderIcon -Icon "$IconfullPath" -Path "$subfolder"
				attrib +R +S "$subfolder"
				#Rename-Item "$subfolder" "$subfolder _tmp"
				#Start-Sleep -s 1
				#Rename-Item "$subfolder _tmp" "$subfolder"
			}
		}else{
			Write-Host "$subfolder is not blank"
			#Out-File -FilePath $folder\log_all.txt -Append -InputObject "$subfolder is not blank"
			#Remove-Item $subfolder\DESKTOP.INI -Force -errorAction SilentlyContinue -errorVariable errors
			if ([System.IO.File]::Exists("$subfolder\$INIFile")){
				Remove-SetIcon -Path "$subfolder"
				attrib -R -S "$subfolder"
				#Rename-Item "$subfolder" "$subfolder _tmp"
				#Start-Sleep -s 1
				#Rename-Item "$subfolder _tmp" "$subfolder"
			}
		}
		}
		#TASKKILL /F /FI "WINDOWTITLE eq $fullpath" /IM explorer.exe
			
		#Stop-Process -ProcessName explorer
		#Get-Process explorer | Foreach-Object { Write-Host $_.MainWindowTitle} 
		#get the active window, then kill it
		$activeFolder=Split-Path -Leaf $path
		Write-Host "activeFolder:$activeFolder"
		Get-Process explorer | Where-Object {
		if($_.MainWindowTitle -eq "$activeFolder")
		{
			$_.CloseMainWindow() 
			Write-Host $_.MainWindowTitle
		}
			else{
			#do nothing
			}
		}
		taskkill /fi “imagename eq explorer.exe” /f	
		#gps explorer|%{$_.MainWindowTitle}
		if($changeType -eq "Created"){
			START explorer.exe  "/select, $fullpath"
		}else{
			START explorer.exe "$path"
			}
		}
}

#5. "Subscribe" to any events where our file/folder is created or deleted
Register-ObjectEvent -InputObject $FileSystemWatcher -EventName "Created" -Action $ChangeIconAction -MessageData $pso
Register-ObjectEvent -InputObject $FileSystemWatcher -EventName "Deleted" -Action $ChangeIconAction -MessageData $pso
#Register-ObjectEvent -InputObject $FileSystemWatcher -EventName "Error" -Action $EmailNotifyAction -MessageData $pso

# To stop the monitoring, run the following commands: 
# Unregister-Event Created.id

