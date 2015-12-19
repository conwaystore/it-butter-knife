# it-butter-knife
> No es un Swiss-Army-Knife pero sí muy esencial para todo profesional IT de NT

Compartimos nuestros snippets a la comunidad. Tal vez encuentre uno util para usted :metal:

## Snippets
### MapDriveToSMBPath.vbs
> Mapea unidad logica a una remota (via SMB)

#### Variables

##### strRemotePath
Inicializacion: `Requerida` <br />
Tipo: `string`

##### strDriveLabel
Inicializacion: `Requerida` <br />
Tipo: `string`

Note que en el siguiente snippet hemos utilizado la funcion **nextDriveLabel**.

 ```vbnet
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
 ```

## Licencia
MIT © 2015
