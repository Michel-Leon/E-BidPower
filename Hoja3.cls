'===================Hoja  glosa=======================
Public Sub worksheet_Activate()
    
    With Me.OLEObjects("Tipocelda").Object
        .clear
        .AddItem "Primaria"
        .AddItem "Segundaria"
    End With  
     
    With Me.OLEObjects("C1").Object
        .clear
        .AddItem "Condiciones Generales"
    End With  
    
    With Me.OLEObjects("C2").Object
        .clear
        .AddItem "Caracteristicas Generales"
    End With 
    With Me.OLEObjects("C3").Object
        .clear
        .AddItem "IM"
        .AddItem "QM"
        .AddItem "DM1-A"
        .AddItem "GAM-0"
        .AddItem "GAM-2"
        .AddItem "GAM-3"
        .AddItem "GBC"
        .AddItem "NSM-0"
        .AddItem "NSM-1"
        .AddItem "NSM-2"
        .AddItem "NSM-3"
    End With
    With Me.OLEObjects("C4").Object
        .clear
        .AddItem "IM"
        .AddItem "QM"
        .AddItem "DM1-A"
        .AddItem "GAM-0"
        .AddItem "GAM-2"
        .AddItem "GAM-3"
        .AddItem "GBC"
        .AddItem "NSM-0"
        .AddItem "NSM-1"
        .AddItem "NSM-2"
        .AddItem "NSM-3"
    End With
    With Me.OLEObjects("C5").Object
        .clear
        .AddItem "IM"
        .AddItem "QM"
        .AddItem "DM1-A"
        .AddItem "GAM-0"
        .AddItem "GAM-2"
        .AddItem "GAM-3"
        .AddItem "GBC"
        .AddItem "NSM-0"
        .AddItem "NSM-1"
        .AddItem "NSM-2"
        .AddItem "NSM-3"
    End With
    With Me.OLEObjects("C6").Object
        .clear
        .AddItem "IM"
        .AddItem "QM"
        .AddItem "DM1-A"
        .AddItem "GAM-0"
        .AddItem "GAM-2"
        .AddItem "GAM-3"
        .AddItem "GBC"
        .AddItem "NSM-0"
        .AddItem "NSM-1"
        .AddItem "NSM-2"
        .AddItem "NSM-3"
    End With
End Sub

private sub Tipocelda_change()
    dim celda as string
    celda = Me.OLEObjects("Tipocelda").Object.Value
    With Me.OLEObjects("select2").Object
        .clear
        select case celda
            case "Primaria"
                .AddItem "24Kv"
                .AddItem "36Kv"
            case "Segundaria"
                .AddItem "24Kv"
                .AddItem "36Kv"
        end select
    End With
End sub
private sub select2_change()
    dim selecccion as string
    seleccion = Me.OLEObjects("select2").Object.Value
    'aqui va el codigo para cada seleccion
    With Me.OLEObjects("select3").Object
        .clear
        select case seleccion
            case "24Kv"
                .AddItem "AIS"
                .AddItem "2SIS"
            case "36Kv"
                .AddItem "AIS"
        end select
    End With
end sub
private sub select3_change()
    dim seleccion as string
    seleccion = Me.OLEObjects("select3").Object.Value
    With Me.OLEObjects("select4").Object
        .clear
        select case seleccion
            case "AIS"
                .AddItem "SM6"
                .AddItem "SM AIRSET"
            case "2SIS"
                .AddItem "PREMSET"
        end select
    End With
end sub

' FUNCION PARA TOMAR VALOR DE LOS COMBOBOX Y GUARDARLOS EN VARIABLES PUBLICAS

Private Sub btnCargarDocumentacion_Click()
    Dim valorC1 As String
    
    valorC1 = Me.OLEObjects("C1").Object.Value
    
    Select Case valorC1
        Case "Condiciones Generales"
            posicion1_1 = BuscarEnBaseGlosa("CG_1") & vbCrLf & BuscarEnBaseGlosa("CG1_2")
            'posicion1_2 = BuscarEnBaseGlosa("CG1_2")
            posicion1_3 = BuscarEnBaseGlosa("CG1_3")
            
        Case Else
            MsgBox "Seleccione una opción válida.", vbExclamation
            Exit Sub
    End Select
    
    Call ActualizarContenidoWord
End Sub 
