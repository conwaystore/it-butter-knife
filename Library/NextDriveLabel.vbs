Function nextDriveLabel()
	Dim objDict, _
			objWMI, _
			objNet, _
			objDisk, _
			colDisks
	Dim strDriveLabel, _
			i

	Set objDict = CreateObject("Scripting.Dictionary")
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}\\.\root\cimv2")
	Set colDisks = objWMI.ExecQuery("Select * from Win32_LogicalDisk")

	For Each objDisk In colDisks
		objDict.Add objDisk.DeviceID, objDisk.DeviceID
	Next

	For i = 67 To 90
		strDriveLabel = Chr(i) & ":"

		If Not objDict.Exists(strDriveLabel) Then
			nextDriveLabel = strDriveLabel
		End If
	Next
End Function
