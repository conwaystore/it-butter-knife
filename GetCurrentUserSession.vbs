Option Explicit

Dim strComputer
Dim strUserName
Dim strUserPassword
Dim objSWbemLocator, objWMIService

strComputer = "."

' CONSTANTS
Const SWBENSECURITY_AUTHENTICATION_LEVEL = 6
Const SWBENSECURITY_IMPERSONATION_LEVEL = 3

If strUserName = "" Then
	Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
	Set objWMIService = objSWbemLocator.ConnectServer _
		(strComputer, "root\cimv2", strUserName, strUserPassword)

	objWMIService.Security_.ImpersonationLevel = SWBENSECURITY_IMPERSONATION_LEVEL
	objWMIService.Security_.AuthenticationLevel = SWBENSECURITY_AUTHENTICATION_LEVEL
Else
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\ " _
						& strComputer & "\root\cimv")
End If


Dim colSessions
Dim colLists
Dim objSession
Dim objItem

Set colSessions = objWMIService.ExecQuery("Select * From Win32_LogonSession Where LogonType=2")

For Each objSession In colSessions
	Set colLists = objWMIService.ExecQuery("Associators Of " _
		& "{Win32_LogonSession.LogonId=" & objSession.LogonId & "} " _
		& "Where AssocClass=Win32_LoggedOnUser Role=Dependent")

	For Each objItem  In colLists
		strUserName = objItem.Name
	Next
Next

Wscript.Echo strUserName