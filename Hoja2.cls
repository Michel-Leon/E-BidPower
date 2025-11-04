'==============HOJA PROYECTO================
public sub  Worksheet_Activate()
    with Me.OLEObjects("Plantillas").Object
        .clear
        .additem "Plantilla 1"
        .additem "Plantilla 2"
        .additem "Plantilla 3"
        .additem "Plantilla 4"
    end with
End Sub
