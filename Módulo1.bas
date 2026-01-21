Attribute VB_Name = "ModuloPublic"

Public posicion1_1 as string, posicion1_2 as string, posicion1_3 as string

Public Sub CargarPlantillaEnGlosa()
    Dim RutaPlantilla As String
    Dim Plantilla As String
    Dim ws As Worksheet
    Dim wsProyecto As Worksheet
    Dim objWord As OLEObject
    Dim ancho As Double, alto As Double
    Dim wsBaseGlosa As Worksheet
    
    Set wsProyecto = ThisWorkbook.Sheets("Proyecto")
    Set ws = ThisWorkbook.Sheets("Glosa")
    Set wsBaseGlosa = ThisWorkbook.Sheets("Glosa_MT24")
    
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
            RutaPlantilla = "C:\Users\sleon\Documents\0. BASES DE DATOS\PROKIT\0.0 Plantillas\GLO-PLA-1.docx"
        Case "Plantilla 2"
            RutaPlantilla = "C:\Users\sleon\Documents\0. BASES DE DATOS\PROKIT\0.0 Plantillas\GLO-PLA-2.docx"
        Case "Plantilla 3"
            RutaPlantilla = "C:\Users\sleon\Documents\0. BASES DE DATOS\PROKIT\0.0 Plantillas\GLO-PLA-3.docx"
        Case "Plantilla 4"
            RutaPlantilla = "C:\Users\sleon\Documents\0. BASES DE DATOS\PROKIT\0.0 Plantillas\GLO-PLA-4.docx"
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
    NombreProyecto = wsProyecto.Range("DZ55").Value
    NombreEmpresa = wsProyecto.Range("DZ64").Value
    
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
' FUNCIÓN PARA BUSCAR EN LA BASE DE DATOS Y OBTENER VALORES
Public Function BuscarEnBaseGlosa(nombreCelda As String) As String
    Dim valor As String
    
    On Error Resume Next
    valor = ThisWorkbook.Names(nombreCelda).RefersToRange.Value
    On Error GoTo 0
    
    If valor <> "" Then
        BuscarEnBaseGlosa = valor
    Else
        BuscarEnBaseGlosa = ""
        MsgBox "No se encontró la celda con nombre: " & nombreCelda, vbExclamation
    End If
End Function
' FUNCIÓN MEJORADA PARA ACTUALIZAR MARCADOR SIN ELIMINARLO
Private Sub ActualizarMarcador(wordDoc As Object, nombreMarcador As String, textoNuevo As String, nivelTitulo As Integer)
    Dim rng As Object
    Dim rangoInicio As Long
    Dim rangoFin As Long
    Dim textoSinNumero As String
    
    If wordDoc.Bookmarks.Exists(nombreMarcador) Then
        Set rng = wordDoc.Bookmarks(nombreMarcador).Range

        'Quitar numeracion manual si existe (ej: "1.1 descripcion" -> "descripcion")
        TextoSinNumero = QuitarNumeracion(textoNuevo)
        rng.Text = TextoSinNumero
        rangoInicio = rng.Start
        rangoFin = rng.Start + Len(TextoSinNumero)

        'Aplicar estilo con numeracion automatica
        Select case nivelTitulo
            Case 1
                wordDoc.Range(rangoInicio, rangoFin).Style = "Título 1"
            Case 2
                wordDoc.Range(rangoInicio, rangoFin).Style = "Título 2"
            Case 3
                wordDoc.Range(rangoInicio, rangoFin).Style = "Título 3"
            Case Else
                wordDoc.Range(rangoInicio, rangoFin).Style = "Normal"    
        End Select
        wordDoc.Bookmarks.Add nombreMarcador, wordDoc.Range(rangoInicio, rangoFin)
    End If
End Sub
' FUNCIÓN AUXILIAR PARA QUITAR NUMERACIÓN MANUAL
Private Function QuitarNumeracion(texto As String) As String
    Dim i As Integer
    Dim resultado As String
    
    resultado = Trim(texto)
    
    ' Buscar donde termina la numeración (después de números y puntos)
    For i = 1 To Len(resultado)
        If Not (Mid(resultado, i, 1) Like "#" Or Mid(resultado, i, 1) = "." Or Mid(resultado, i, 1) = " ") Then
            resultado = Trim(Mid(resultado, i))
            Exit For
        End If
    Next i
    
    QuitarNumeracion = resultado
End Function
' FUNCIÓN PARA INSERTAR IMAGEN EN WORD DESPUÉS DE UN MARCADOR
Private Sub InsertarImagenDespuesDeMarcador(wordDoc As Object, nombreMarcador As String, nombreImagen As String)
    Dim wsBaseGlosa As Worksheet
    Dim shp As Shape
    Dim rng As Object
    
    On Error GoTo ErrorImagen
    
    Set wsBaseGlosa = ThisWorkbook.Sheets("Glosa_MT24")
    
    ' Buscar la imagen por su nombre
    Set shp = wsBaseGlosa.Shapes(nombreImagen)
    
    ' Copiar la imagen
    shp.Copy
    
    ' Ir al final del marcador en Word y pegar
    If wordDoc.Bookmarks.Exists(nombreMarcador) Then
        Set rng = wordDoc.Bookmarks(nombreMarcador).Range
        rng.Collapse Direction:=0  ' wdCollapseEnd
        rng.InsertParagraphAfter
        rng.Collapse Direction:=0
        rng.Paste
        ' Obtener la imagen recien pegada
        Set imgWord = wordDoc.InlineShapes(wordDoc.InlineShapes.Count)

        '====== Ajustar Tamaño =====
        imgWord.Width = shp.Width
        imgWord.Height = shp.Height
        rng.paragraphFormat.Alignment = 1  ' 0=Izquierda, 1=Centro, 2=Derecha
    End If
    
    Exit Sub
    
ErrorImagen:
    MsgBox "No se encontró la imagen: " & nombreImagen, vbExclamation
End Sub
'FUNCIÓN PRINCIPAL PARA ACTUALIZAR CONTENIDO EN WORD
Public Sub ActualizarContenidoWord()
    Dim ws As Worksheet
    Dim objWord As OLEObject
    Dim wordApp As Object
    Dim wordDoc As Object
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Glosa")
    
    On Error Resume Next
    Set objWord = ws.OLEObjects("objWord")
    On Error GoTo ErrorHandler
    
    If objWord Is Nothing Then
        MsgBox "Primero debe cargar una plantilla.", vbExclamation
        Exit Sub
    End If
    
    objWord.Activate
    Set wordApp = objWord.Object.Application
    Set wordDoc = wordApp.ActiveDocument
    
    ' Actualizar con estilos de título
    'Nivel 1 = Título principal (1., 2., 3.)
    'Nivel 2 = Subtítulo (1.1, 1.2, 2.1)
    'Nivel 0 = Texto normal


    Call ActualizarMarcador(wordDoc, "Posicion1_1", posicion1_1, 1)
    ' Insertar imagen después de Posicion1_1
    Call InsertarImagenDespuesDeMarcador(wordDoc, "Posicion1_1", "SM6 24Kv")
    Call ActualizarMarcador(wordDoc, "Posicion1_2", posicion1_2, 0)
    Call ActualizarMarcador(wordDoc, "Posicion1_3", posicion1_3, 0)
    
    MsgBox "Contenido actualizado correctamente.", vbInformation
    
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Exit Sub

    ErrorHandler:
        MsgBox "Error: " & Err.Description, vbExclamation
        Set wordDoc = Nothing
        Set wordApp = Nothing
End Sub
'=================== MÓDULO - CÓDIGO DE DIAGNÓSTICO ===================
Public Sub DiagnosticarProblema()
    Dim wsBaseGlosa As Worksheet
    Dim wsGlosa As Worksheet
    Dim objWord As OLEObject
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim bm As Object
    Dim celda As Range
    Dim msg As String
    
    Set wsBaseGlosa = ThisWorkbook.Sheets("Glosa_MT24")
    Set wsGlosa = ThisWorkbook.Sheets("Glosa")
    
    msg = "=== DIAGNÓSTICO ===" & vbCrLf & vbCrLf
    
    ' 1. BUSCAR LOS CÓDIGOS EN LA HOJA
    msg = msg & "1. BÚSQUEDA EN GLOSA_MT24:" & vbCrLf
    
    Set celda = wsBaseGlosa.Cells.Find(What:="CG_1", LookIn:=xlValues, LookAt:=xlWhole)
    If Not celda Is Nothing Then
        msg = msg & "   CG_1 encontrado en: " & celda.Address & " | Valor al lado: " & celda.Offset(0, 1).Value & vbCrLf
    Else
        msg = msg & "   CG_1 NO ENCONTRADO" & vbCrLf
    End If
    
    Set celda = wsBaseGlosa.Cells.Find(What:="CG1_2", LookIn:=xlValues, LookAt:=xlWhole)
    If Not celda Is Nothing Then
        msg = msg & "   CG1_2 encontrado en: " & celda.Address & " | Valor al lado: " & celda.Offset(0, 1).Value & vbCrLf
    Else
        msg = msg & "   CG1_2 NO ENCONTRADO" & vbCrLf
    End If
    
    Set celda = wsBaseGlosa.Cells.Find(What:="CG1_3", LookIn:=xlValues, LookAt:=xlWhole)
    If Not celda Is Nothing Then
        msg = msg & "   CG1_3 encontrado en: " & celda.Address & " | Valor al lado: " & celda.Offset(0, 1).Value & vbCrLf
    Else
        msg = msg & "   CG1_3 NO ENCONTRADO" & vbCrLf
    End If
    
    ' 2. VERIFICAR MARCADORES EN WORD
    msg = msg & vbCrLf & "2. MARCADORES EN WORD:" & vbCrLf
    
    On Error Resume Next
    Set objWord = wsGlosa.OLEObjects("objWord")
    On Error GoTo 0
    
    If objWord Is Nothing Then
        msg = msg & "   NO HAY DOCUMENTO WORD CARGADO" & vbCrLf
    Else
        objWord.Activate
        Set wordApp = objWord.Object.Application
        Set wordDoc = wordApp.ActiveDocument
        
        For Each bm In wordDoc.Bookmarks
            msg = msg & "   - " & bm.Name & vbCrLf
        Next bm
        
        Set wordDoc = Nothing
        Set wordApp = Nothing
    End If
    
    MsgBox msg, vbInformation, "Resultado del Diagnóstico"
End Sub