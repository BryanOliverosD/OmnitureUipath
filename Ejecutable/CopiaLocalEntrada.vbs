Option Explicit
Dim objFSO   'objeto del tipo FileSystemObject para manejar el fichero
Dim sobreescribir         	'variable que contendra true o false.

On Error Resume Next

sobreescribir = True    	'En mi caso no quiero que sobreescriba el fichero si existe.
Set objFSO = CreateObject("Scripting.FileSystemObject")

'comprobar si existe el directorio c:\destino existe y si es asi, copiar el fichero.
If (objFSO.FolderExists("Y:\root\Omniture\Entrada")) Then
	objFSO.CopyFile  "Y:\root\Omniture\Entrada\*.*", "C:\Users\jcarrillo\Desktop\Omniture\Entrada\", sobreescribir
	objFSO.DeleteFile "Y:\root\Omniture\Entrada\*.*"
End If
On Error Goto 0