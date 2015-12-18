Dim strComputer, _
		objWMI, _
		colDisks, _
		objDict
strComputer = "."

Set objWMI = GetObject("winmgmts:\\" _
														& strComputer _
														& "\root\cimv2")
Set colDisks = objWMI.ExecQuery ("Select * " _
																			& "From Win32_LogicalDisk")

Set objDict = CreateObject("Scripting.Dictionary")

For Each objDrive In colDisks
	objDict.Add objDrive.DeviceID, objDrive.ProviderName
Next

If objDict.Exists("X:") Then
	Wscript.Echo "Exists"
Else
	Wscript.Echo "No Exists"
End If
