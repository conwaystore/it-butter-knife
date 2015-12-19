# Library

Library incluye funciones y rutinas.

## Funciones
### nexDriveLabel()
Retorna: `string`

Automaticamente genera la siguiente etiqueta **sin uso** para discos logicos.

##### Ejemplo
```vbnet
Option Explicit

Dim strDriveLabel
strDriveLabel = nexDriveLabel

' Letra de suerte!
Wscript.Echo strDriveLabel
```
