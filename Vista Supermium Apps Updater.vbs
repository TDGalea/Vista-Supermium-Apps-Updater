On Error Resume Next

Set oShell   = WScript.CreateObject("WScript.Shell")
Set fso      = CreateObject("Scripting.FileSystemObject")

Function CreateShortcut
	sLinkFile = appdata + "\Microsoft\Windows\Start Menu\Programs\Supermium Apps\" + name + ".lnk"
	Set oLink = oShell.CreateShortcut(sLinkFile)
		oLink.TargetPath   = "C:\Program Files\Supermium\chrome_proxy.exe"
		oLink.Arguments    = "--profile-directory=""" + profile + """ --app-id=" + appid
		oLink.IconLocation = localappdata + "\Supermium\User Data\" + profile + "\Web Applications\_crx_" + appid + "\" + name + ".ico"
	oLink.Save
End Function

localappdata = oShell.ExpandEnvironmentStrings("%localappdata%")
appdata      = oShell.ExpandEnvironmentStrings("%appdata%")

Set userData = fso.GetFolder(localappdata + "\Supermium\User Data\")

appList     = ""
prevProfile = ""

For Each profFolder in userData.SubFolders
	profile = profFolder.Name
	Set manifestFolder = fso.GetFolder(localappdata + "\Supermium\User Data\" + profile + "\Web Applications\Manifest Resources")
	
	If fso.FolderExists(manifestFolder) Then
		
		For Each manFolder in manifestFolder.SubFolders
			appid = manFolder.Name
			iconFolderPath = localappdata + "\Supermium\User Data\" + profile + "\Web Applications\_crx_" + appid
			
			If fso.FolderExists(iconFolderPath) Then
				Set iconFolder = fso.GetFolder(iconFolderPath)
				For Each file In iconFolder.Files
					If LCase(fso.GetExtensionName(file.Name)) = "ico" Then
						name = fso.GetBaseName(file.Name)
						CreateShortcut
						If profile = prevProfile Then
							appList = appList + ", " + name
						Else
							appList = appList + vbCrLf + vbCrLf + "Profile: " + profile + vbCrLf + "Apps: " + name
						End If
						prevProfile = profile
						Exit For
					End If
				Next
			End If
		Next
	End If
Next


x = msgbox("The following WebApp shortcuts have been created:" + appList,64,"Supermium Vista WebApps Updater")
