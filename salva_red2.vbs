Dim fso, objNet, Carpeta, nivel, cont_archivos, cont1, Final, Start
Const ForReading = 1, ForWriting = 2, ForAppending = 8

'********************************************************************************************************************************************
Class File
	'************************************************************************
	'Declaración de propiedades
	Public Name
	Public DateLastModified
	Public size
	Public Path

	'************************************************************************
	'Declaración de métodos

	Sub Copy (Destination, overwrite)
		
	End Sub
	'****************************************
	Sub Move (Destination)
		
	End Sub
	'****************************************
	Sub Delete (force)
		
	End Sub

End Class

'********************************************************************************************************************************************	
Class Folder

	'************************************************************************
	'Declaración de propiedades
	Public Name
	Public DateLastModified
	Public size
	Public Path
	Public Files
	Public Subfolders
	'************************************************************************
	'Declaración de métodos
	
	Private Sub Class_Initialize	' Setup Initialize event.
	
		Set Files = CreateObject("Scripting.Dictionary")
		Set Subfolders = CreateObject("Scripting.Dictionary")
	
	End Sub
	
	
End Class
'********************************************************************************************************************************************	
Sub Escribir_estructura(ByRef Nombre)
	Dim ArregloFiles, ArregloCarpetas

	MyFile_salida.writeline("C\" & CStr(nivel-1) &"\" & Nombre.Name & "\" & Nombre.DateLastModified & "\" & Nombre.size)
	ArregloFiles = Nombre.Files.Items
	For each f in arregloFiles
		MyFile_salida.Writeline( "A\" & CStr(nivel) & "\" & f.Name & "\" & f.DateLastModified & "\" & f.size)	
	next
	ArregloCarpetas = Nombre.Subfolders.Items
	For each C in ArregloCarpetas
		nivel = nivel + 1	
		Escribir_estructura( Nombre.Subfolders.Item(C.Name))
		nivel = nivel - 1
	next
	
End Sub
'***************************************************************************************************************************************
Sub Anticomparacion(Virtual,Real,bandera)
	Dim ArregloFiles, ArregloCarpetas
	
	ArregloFiles = Virtual.Files.Items
	if bandera = true then
		Set folder_origen = fso.GetFolder(Real)
		Set	archivos_origen = folder_origen.Files
		Set sub_carpetas_origen = folder_origen.Subfolders
	End if
	For each f in arregloFiles
		if bandera = true then
			If  Not (fso.FileExists(folder_origen.Path & "\" & f.Name)) and (DateDiff("m",CDate(f.DateLastModified),Now) > 3) then
					fso.DeleteFile Virtual.Path & "\" & f.Name,True
					Virtual.Files.Remove(f.Name)
			end If
		Else
			if DateDiff("m",CDate(f.DateLastModified),Now) > 3 then
				fso.DeleteFile Virtual.Path & "\" & f.Name,True
				Virtual.Files.Remove(f.Name)
			End If
		End if
		If cont_archivos = 20 then
			cont_archivos = 0
			cont1 = cont1 + 1
			Final =Timer
			Set MyFile = fso.OpenTextFile("debug.txt",ForAppending,True)
			MyFile.WriteLine(CStr(Now) & ":    " & "SubIteración #" & CStr(cont1) &": " & (Final - Start) & "s")
			MyFile.Close
			Wscript.sleep 500
			Start = Timer
		Else
			Cont_archivos = Cont_archivos + 1
		End If
	Next
	ArregloCarpetas = Virtual.Subfolders.Items
	For each C in ArregloCarpetas
		If bandera = true then
			If fso.FolderExists(folder_origen.path & "\" & C.Name) then
				Call Anticomparacion(Virtual.Subfolders.Item(C.Name),sub_carpetas_origen.Item(C.Name).Path, true)
			else
				if C.Subfolders.Count = 0 and C.Files.Count = 0 then
					virtual.Subfolders.Remove(C.Name)
					fso.DeleteFolder Virtual.Path & "\" & C.Name,True
				Else
					Call Anticomparacion(Virtual.Subfolders.Item(C.Name),"Cualquier_cosa", false)
				End If
			End if
		Else
			if C.Subfolders.Count = 0 and C.Files.Count = 0 then
				virtual.Subfolders.Remove(C.Name)
				fso.DeleteFolder Virtual.Path & "\" & C.Name,True
			Else
				Call Anticomparacion(Virtual.Subfolders.Item(C.Name),"Cualquier_cosa", false)
			End If
		End If
	next
	Set sub_carpetas_origen = nothing
	Set archivos_origen = nothing
	Set folder_origen = nothing

End Sub
'***************************************************************************************************************************************
Sub Estructura_carpetas(ByVal Nombre, ByVal NombreVirtual)

	Set folder_origen = fso.GetFolder(Nombre)
	Set archivos_origen = folder_origen.Files
	Set sub_carpetas_origen = folder_origen.Subfolders
	for each f in archivos_origen 
		Set Archivo = New File
		Archivo.Name = f.Name
		Archivo.DateLastModified = f.DateLastModified
		Archivo.size = f.size
		NombreVirtual.Files.Add Archivo.Name, Archivo
		Set Archivo = Nothing
	next
	for each C in sub_carpetas_origen
		Set subcarpeta = New Folder
		subcarpeta.Name = C.Name
		subcarpeta.Path = C.Path
		subcarpeta.size = C.size
		subcarpeta.DateLastModified = C.DateLastModified
		NombreVirtual.Subfolders.Add subcarpeta.Name, subcarpeta
		Call Estructura_carpetas(C.path,NombreVirtual.Subfolders.Item(subcarpeta.Name))
		Set subcarpeta = Nothing
	Next
	Set sub_carpetas_origen = nothing
	Set archivos_origen = nothing
	Set folder_origen = nothing
	
End Sub
'***************************************************************************************************************************************
Sub Comparacion(ByRef Virtual,ByVal Real)

	Set folder_origen = fso.GetFolder(Real)
	Set archivos_origen = folder_origen.Files
	Set sub_carpetas_origen = folder_origen.Subfolders
	for each f in archivos_origen
		copiar = true
		If  Virtual.Files.Exists(f.Name) then 
			If Virtual.Files.Item(f.Name).DateLastModified = archivos_origen.Item(f.Name).DateLastModified then
				Copiar = false
			Else
				Virtual.Files.Remove f.Name
			End If
		End If
		if copiar then
			archivos_origen.Item(f.Name).Copy(Virtual.Path & "\" & f.Name)
			Set Archivo = New File
			Archivo.Name = f.Name
			Archivo.DateLastModified = f.DateLastModified
			Archivo.size = f.size
			Virtual.Files.Add Archivo.Name, Archivo
			Set Archivo = Nothing
		End if
		If cont_archivos = 20 then
			cont_archivos = 0
			cont1 = cont1 + 1
			Final =Timer
			Set MyFile = fso.OpenTextFile("debug.txt",ForAppending,True)
			MyFile.WriteLine(CStr(Now) & ":    " & "SubIteración #" & CStr(cont1) &": " & (Final - Start) & "s")
			MyFile.Close
			Wscript.sleep 500
			Start = Timer
		Else
			Cont_archivos = Cont_archivos + 1
		End If
	next
	for each C in sub_carpetas_origen
		if Not Virtual.Subfolders.Exists(C.Name) then
			Set subcarpeta = New Folder
			subcarpeta.Name = C.Name
			subcarpeta.size = C.size
			subcarpeta.DateLastModified = C.DateLastModified
			subcarpeta.Path = Virtual.Path & "\" & C.Name
			Virtual.Subfolders.Add subcarpeta.Name, subcarpeta
			fso.CreateFolder virtual.subfolders.Item(subcarpeta.Name).Path
			Set subcarpeta = Nothing
		End if
		Call Comparacion(Virtual.Subfolders.Item(C.Name),C.path)
	next
		
End Sub
'***************************************************************************************************************************************
Sub Conectar(ByRef valor,ByVal valor1,ByVal valor2,ByVal valor3,ByVal valor4)

	Inicio = Timer
	Set objShell = CreateObject("WScript.Shell")
	Do
		ret_val = objShell.Run("ping " & valor1, 0, True)
		if ret_val = 0 then
			i = asc(valor)
			Do
				if Not fso.DriveExists(chr(i)) then Exit Do
				i = i - 1
			Loop Until i = 67
			if Not i = 67 then
				'msgbox "i:" & chr(i)
				objNet.MapNetworkDrive chr(i) & ":", "\\" & valor1 & "\" & valor2, False, valor3, valor4
				valor = chr(i) & ":"
				Exit Do
			End If
			cadena = "No hay letra de unidad disponible para conectar la unidad de red"
			cod = 1002
		else
			cadena = "No se puede hacer contacto con el servidor " & chr(34) & valor1 & chr(34)
			cod = 1001
		end if
		Final = Timer
		
		Set objShell = Nothing
		
		Wscript.sleep 5000
		
		if (Final - Inicio) > 60 then '900
			a = error(cadena,cod)
			Inicio = Timer
		End if
	Loop
	
End Sub
'***************************************************************************************************************************************
'En este procedimiento debiera buscarse la manera de chequear que no hay ningún archivo abierto en el disco mapeado antes de cerrarlo.

Sub Desconectar(ByVal Share)

	objNet.RemoveNetworkDrive Share

End Sub
'***************************************************************************************************************************************
Function error(ByRef cadena,codigo)

	error = msgbox("Error: " & cadena & Chr(13) & "Si oprime Cancelar el sistema de salva automática dejará de funcionar",53,"Salva Automática")
	MyFile.WriteLine(codigo & ":" &cadena)
	if error = 2 then
		msgbox "Para resolver el problema póngase en contacto con el administrador del sistema" & Chr(13) & "comunicándole el siguiente codigo de error: " & chr(34) & codigo & chr(34),0
		wscript.quit
	end if
End Function
'***************************************************************************************************************************************
'***************************************************************************************************************************************
'***************************************************************************************************************************************

'Aquí comienza main

	'Inicialización de objetos del file system y el objeto de red; se borra el archivo "terminado" y se crea el semáforo
	'Posteriormente se leen los datos de trabajo desde el archivo "settings.txt"

	StartTime = Timer
	Set objNet = WScript.CreateObject("WScript.Network")
	Set fso = CreateObject("Scripting.FileSystemObject")
	if fso.FileExists("terminado") then fso.DeleteFile "terminado", True
	if not fso.FileExists("semaforo") then fso.CreateTextFile("semaforo")
	Set MyFile = fso.OpenTextFile("debug.txt",ForWriting,True)
	Do While Not fso.FileExists("settings.txt")
		cadena = "No se encuentra el archivo de inicialización" & chr(34) & "settings.txt" & chr(34)
		a = error(cadena, 1003)
		Wscript.sleep 5000
	Loop
	Set File_Settings = fso.OpenTextFile("settings.txt", forReading, true, 0)
		Do
			temporal = Left(File_Settings.ReadLine,1)
		Loop while (temporal = "'") or (temporal = " ")
		Servidor = File_Settings.ReadLine
		Carpeta_Salva = File_Settings.ReadLine
		Usuario = File_Settings.ReadLine
		Clave = File_Settings.ReadLine
		DriveSalva = File_Settings.ReadLine
		Origen = File_Settings.ReadLine
		Tiempo = File_Settings.ReadLine
		File_Settings.close
	Set File_Settings = Nothing
	Do While Not fso.FolderExists(Origen)
		cadena = "No se encuentra la carpeta a salvar " & chr(34) & Origen & chr(34)
		a = error(cadena,1004)
		'Wscript.sleep 5000
	Loop
	FinalTime = Timer
	MyFile.WriteLine(CStr(Now) & ":    " & "Tiempo de inicialización del programa: " & CStr(FinalTime - StartTime) & "s")
	
	'***************************************************************************************************************************************
	'Aquí se lee la estructura de las carpetas desde el sistema de archivo remoto para el objeto carpeta virtual en memoria
	
	'Wscript.sleep 300000
	StartTime = Timer
	Set Carpeta = New Folder
	Call Conectar(DriveSalva,Servidor,Carpeta_Salva,Usuario,Clave)
	Carpeta.Name = CStr(fso.GetDriveName(DriveSalva))
	Carpeta.Path = CStr(fso.GetDrive(DriveSalva).Path)
	Call Estructura_carpetas(DriveSalva,Carpeta)
	Call Desconectar(DriveSalva)
	FinalTime = Timer
	MyFile.WriteLine
	MyFile.WriteLine(CStr(Now) & ":    " & "Tiempo de lectura de la estructura de la carpeta: " & CStr(FinalTime - StartTime) & "s")
	MyFile.WriteLine
	MyFile.Close
	
	'***************************************************************************************************************************************
	'Aquí comienza el ciclo de comparación entre lo que hay en el sistema de archivos y lo que hay en el objeto carpeta virtual
	cont = 0
	Call Conectar(DriveSalva,Servidor,Carpeta_Salva,Usuario,Clave)		
		Do
			StartTime = Timer
			Start = Timer
			cont1 = 0
			cont_archivos = 0
			Call Comparacion(Carpeta,Origen)
			Call Anticomparacion(Carpeta,Origen,true)
			cont = cont + 1
			FinalTime = Timer
			Set MyFile = fso.OpenTextFile("debug.txt",ForAppending,True)
			MyFile.WriteLine(CStr(Now) & ":    " & "Iteración #" & CStr(cont) &": " & (FinalTime - StartTime) & "s")
			MyFile.Close
		'wscript.sleep 5000
		'Call Conectar(DriveSalva,Servidor,Left(Carpeta_salva,InStr(Carpeta_Salva,"\")-1),Usuario,Clave)
		'	fso.CopyFile "debug.txt",DriveSalva & "\" & Mid(Carpeta_salva,InStrRev(Carpeta_Salva,"\")+1) & "_debug.txt",True
		'Call Desconectar(DriveSalva)
		'Wscript.sleep CLng(Tiempo)
		Loop While fso.FileExists("semaforo")
	Call Desconectar(DriveSalva)
	'***********************************************************************************************************************************
	'Aquí se salva el objeto carpeta virtual para el archivo "salida.txt"
	Set MyFile_salida = fso.OpenTextFile("salida.txt",ForWriting,True,-1)
	nivel = 1
	StartTime = Timer
	Call Escribir_estructura(Carpeta)
	Set MyFile_salida = nothing
	FinalTime = Timer
	Set MyFile = fso.OpenTextFile("debug.txt",ForAppending,True)
	MyFile.WriteLine
	MyFile.WriteLine(CStr(Now) & ":    " & "Tiempo de escritura a disco de la estructura de la carpeta: " & CStr(FinalTime - StartTime) & "s")
	MyFile.Close
	'***************************************************************************************************************************************
	'Limpia de objetos y terminación del programa
	Set Terminado = fso.CreateTextFile("terminado")
	Terminado.Close
	Set Terminado = nothing
	Set MyFile = nothing
	Set fso = nothing
	Set ObjNet = nothing
