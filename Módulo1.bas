Attribute VB_Name = "ModuloPublic"

Public Sub CargarPlantillaEnGlosa()
    Dim RutaPlantilla As String
    Dim Plantilla As String
    Dim ws As Worksheet
    Dim wsProyecto As Worksheet
    Dim objWord As OLEObject
    Dim ancho As Double, alto As Double
    
    Set wsProyecto = ThisWorkbook.Sheets("Proyecto")
    Set ws = ThisWorkbook.Sheets("Glosa")
    
    ' Selección de plantilla desde el ComboBox llamado "Plantillas"
    On Error Resume Next
    Plantilla = wsProyecto.OLEObjects("Plantillas").Object.Value
    On Error GoTo 0
    
    ' Validar que se obtuvo un valor
    If Plantilla = "" Then
        MsgBox "Por favor seleccione una plantilla del ComboBox.", vbExclamation
        Exit Sub
    End If
    
    ' Determinar rutas de las plantillas
    Select Case Plantilla
        Case "Plantilla 1"
            RutaPlantilla = "C:\Users\sleon\Documents\0. BASES DE DATOS\DPI\0.0 Plantillas\GLO-PLA-1.docx"
        Case "Plantilla 2"
            RutaPlantilla = "C:\Users\sleon\Documents\0. BASES DE DATOS\DPI\0.0 Plantillas\GLO-PLA-2.docx"
        Case "Plantilla 3"
            RutaPlantilla = "C:\Users\sleon\Documents\0. BASES DE DATOS\DPI\0.0 Plantillas\GLO-PLA-3.docx"
        Case "Plantilla 4"
            RutaPlantilla = "C:\Users\sleon\Documents\0. BASES DE DATOS\DPI\0.0 Plantillas\GLO-PLA-4.docx"
        Case Else
            MsgBox "NO SE HA SELECCIONADO UNA PLANTILLA VÁLIDA." & vbCrLf & "Valor seleccionado: " & Plantilla, vbExclamation
            Exit Sub
    End Select
    
    ' Verificar que el archivo existe
    If Dir(RutaPlantilla) = "" Then
        MsgBox "La plantilla no existe en la ruta especificada: " & vbCrLf & RutaPlantilla, vbCritical
        Exit Sub
    End If
    
    ' ELIMINAR OBJETO PREVIO
    On Error Resume Next
    ws.OLEObjects("objWord").Delete
    On Error GoTo 0
    
    ' CALCULAR DIMENSIONES (ajusta el rango según necesites)
    ancho = ws.Range("CJ21:HI21").Width
    alto = ws.Range("CJ21:HI177").Height
    
    ' INSERTAR EL OBJETO WORD
    Set objWord = ws.OLEObjects.Add( _
        Filename:=RutaPlantilla, _
        Link:=False, _
        DisplayAsIcon:=False, _
        Left:=ws.Range("CJ21").Left, _
        Top:=ws.Range("CJ21").Top, _
        Width:=ancho, _
        Height:=alto)
    
    objWord.Name = "objWord"
    
    MsgBox "PLANTILLA CARGADA CORRECTAMENTE", vbInformation
    
    ' ACTUALIZAR PORTADA
    ActualizarPortadaWord objWord, wsProyecto
End Sub

Public Sub ActualizarPortadaWord(objWord As OLEObject, wsProyecto As Worksheet)
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim NombreProyecto As String
    Dim NombreEmpresa As String
    
    On Error GoTo ErrorHandler
    
    ' Obtener valores de las celdas
    NombreProyecto = wsProyecto.Range("DZ46").Value
    NombreEmpresa = wsProyecto.Range("DZ51").Value
    
    ' Activar el objeto Word
    objWord.Activate 
    Set wordApp = objWord.Object.Application
    Set wordDoc = wordApp.ActiveDocument
    
    ' Actualizar marcadores
    If wordDoc.Bookmarks.Exists("NombreProyecto") Then
        wordDoc.Bookmarks("NombreProyecto").Range.Text = NombreProyecto
    End If
    
    If wordDoc.Bookmarks.Exists("NombreEmpresa") Then
        wordDoc.Bookmarks("NombreEmpresa").Range.Text = NombreEmpresa
    End If
    
    ' Liberar objetos
    Set wordDoc = Nothing
    Set wordApp = Nothing
    
    Exit Sub

ErrorHandler:
    MsgBox "Error al actualizar la portada: " & Err.Description, vbExclamation
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub