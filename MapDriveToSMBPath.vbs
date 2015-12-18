Option Explicit

Sub Import(vbsSource)
	With CreateObject("Scripting.FileSystemObject")
		Dim strFile

		With CreateObject("Wscript.Shell")
			strFile = .ExpandEnvironmentStrings(vbsSource)
		End With

		ExecuteGlobal .OpenTextFile(.GetAbsolutePathName(strFile)).ReadAll
	End With
End Sub

Import "Library\NextDriveLabel.vbs"
Import "Library\MapNetDrive.vbs"

Dim strDriveLabel
Dim strRemotePath

strDriveLabel = nextDriveLabel
strRemotePath = "\\SOME-IP-ADDRESS\UNC\%USERNAME%"

MapNetDrive strDriveLabel, strRemotePath
