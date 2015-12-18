strRootPath = "C:\SharePub\"

Set dicTSUsers = CreateObject("Scripting.Dictionary")
dicTSUsers.Add "compund2", "compund2"
dicTSUsers.Add "Contabilidad3", "Contabilidad3"
dicTSUsers.Add "contabilidad12", "contabilidad12"
dicTSUsers.Add "controlinvch", "controlinvch"
dicTSUsers.Add "compund3", "compund3"
dicTSUsers.Add "compund1", "compund1"
dicTSUsers.Add "Contabilidad6", "Contabilidad6"
dicTSUsers.Add "bodega2unlp", "bodega2unlp"
dicTSUsers.Add "compras2", "compras2"
dicTSUsers.Add "cexterior", "cexterior"
dicTSUsers.Add "capta1unw", "capta1unw"
dicTSUsers.Add "capta3unw", "capta3unw"
dicTSUsers.Add "proformazl", "proformazl"
dicTSUsers.Add "eregistrounw", "eregistrounw"
dicTSUsers.Add "controlinvunw", "controlinvunw"
dicTSUsers.Add "asanchez", "asanchez"
dicTSUsers.Add "capta2unw", "capta2unw"
dicTSUsers.Add "trafico2", "trafico2"
dicTSUsers.Add "trafico", "trafico"
dicTSUsers.Add "controlinvpb", "controlinvpb"
dicTSUsers.Add "captador2pb", "captador2pb"
dicTSUsers.Add "pedidospacdd", "pedidospacdd"
dicTSUsers.Add "cdm8pa3", "cdm8pa3"
dicTSUsers.Add "captador3pb", "captador3pb"
dicTSUsers.Add "cdm8pa1", "cdm8pa1"
dicTSUsers.Add "captador1pb", "captador1pb"
dicTSUsers.Add "eregistropacdd", "eregistropacdd"
dicTSUsers.Add "Contabilidad7", "Contabilidad7"
dicTSUsers.Add "planeador2pb", "planeador2pb"
dicTSUsers.Add "contabilidad11", "contabilidad11"
dicTSUsers.Add "cexterior2", "cexterior2"
dicTSUsers.Add "compuna2", "compuna2"
dicTSUsers.Add "pedidospbcdd", "pedidospbcdd"
dicTSUsers.Add "eregistropbcdd", "eregistropbcdd"
dicTSUsers.Add "fjeanpierre", "fjeanpierre"
dicTSUsers.Add "obethancourth", "obethancourth"
dicTSUsers.Add "compuna4", "compuna4"
dicTSUsers.Add "eregistrouna", "eregistrouna"
dicTSUsers.Add "zgudino", "zgudino"
dicTSUsers.Add "sap2", "sap2"

With CreateObject("Scripting.FileSystemObject")
	On Error Resume Next

	dictItems = dicTSUsers.Items

	For i = 0 To dicTSUsers.Count - 1
		.CreateFolder(strRootPath & dictItems(i))
	Next
End With
