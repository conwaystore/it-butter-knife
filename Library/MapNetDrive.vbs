Function isMapped(remotePath)
	Dim objWMI, _
			objDisks, _
			colDisk, _
			bolRes

	' En vez de "True" o "False" (VBScript lo evalua de una extrana manera),
	' utilizamos valores enteros

	' 1: True
	' 0: False
	bolRes = 0

	Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
	Set objDisks = objWMI.ExecQuery("Select DeviceID, " _
																& "ProviderName " _
																&	"From Win32_MappedLogicalDisk")

	For Each colDisk In objDisks
		If remotePath = colDisk.ProviderName Then
			bolRes = 1
		End If
	Next

	isMapped = bolRes
End Function

Sub MapNetDrive(label, remotePath)
	On Error Resume Next

	Dim bolIsMapped
	Dim objNet, _
			objWMI, _
			objDisks, _
			objDict, _
			colDisk

	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set objNet = CreateObject("WScript.Network")

	With objWMI
		Set objDisks = .ExecQuery("Select DeviceID, " _
														& "ProviderName " _
														& "From Win32_MappedLogicalDisk")

		For Each colDisk In objDisks
			Dim deviceId, _
					providerName, _
					mapItems

			deviceId = colDisk.DeviceID
			providerName = colDisk.ProviderName

			If providerName = remotePath Then
				objNet.RemoveNetworkDrive deviceId, true
			End If
		Next
	End With

	objNet.MapNetworkDrive label, remotePath
End Sub
