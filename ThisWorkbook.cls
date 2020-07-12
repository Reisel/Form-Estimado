VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    If Hoja81.Shapes("Group 61").Visible = True And Range("BF1") = 1 Then
        Resultado = MsgBox("El Estimado NO ha sido copiado aun en la base de dato" + Chr(13) + _
        "¿Desea salir sin copiar el Estimado?", vbYesNo, "RECORDATORIO...")
        Select Case Resultado
        Case vbYes: GoTo SSS '
        Case vbNo: Cancel = True
        End Select
    End If
SSS:

End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    On Error GoTo DDD
DDD:
End Sub

Private Sub Workbook_Open()
Application.Calculation = xlCalculationAutomatic

    If Hoja81.Range("BF1") <> "1" Then
        Sheets("Estimado Base").Select
        ActiveSheet.Shapes("Picture 47").Select
        Selection.ShapeRange.ZOrder msoSendToBack
    End If
        
    If Hoja81.Range("BF1") = "1" Then
        GoTo asd
    End If

    UserForm4.Show

asd:
    
'    If Sheets("Estimado Base").ProtectContents Or Sheets("Estimado Base"). _
'    ProtectDrawingObjects Or Sheets("Estimado Base").ProtectScenarios Then
'
'        Dim UlFila As String
'        Dim USUARIO As String
'        Dim AA As String
'
'        Call Hoja81.Desproteger
'        ActiveSheet.Shapes("Picture 47").Select ' BOTO ROJO
'        Selection.ShapeRange.ZOrder msoSendToBack
'        Call Hoja81.Proteger
'
'        UlFila = Range("BE4") + 21
'        AA = Range("C" & UlFila)
'
'        If AA = "Reisel Sanchez" Then
'            USUARIO = "SANCHEZREL"
'         End If
'        If AA = "Reisel" Then
'            USUARIO = "REISEL SANCHEZ"
'         End If
'        If AA = "Leyci Colina" Then
'            USUARIO = "COLINALP"
'         End If
'        If AA = "Elvira Alvarado" Then
'            USUARIO = "ALVARADOEGO"
'         End If
'        If AA = "Lorena Gutierrez" Then
'            USUARIO = "GUTIERREZLDX"
'         End If
'        If AA = "Laura Vargas" Then
'            USUARIO = "VARGASLY"
'        End If
'
'        If CreateObject("WScript.Network").UserName = USUARIO Then
'            Call Hoja81.Desproteger
'            ActiveSheet.Shapes("Picture 48").Select   'BOTON VERDE
'            Selection.ShapeRange.ZOrder msoSendToBack
'        End If
'    End If
End Sub


Sub BotonDesbloquear()

    'VERIFICA QUE ESTE BLOQUEADO
 If Sheets("Estimado Base").ProtectContents Or Sheets("Estimado Base"). _
     ProtectDrawingObjects Or Sheets("Estimado Base").ProtectScenarios Then
        
        'VALIDA EL USUARIO
        Dim UlFila As String
        Dim USUARIO As String
        Dim USUARIO2 As String
        Dim AA As String
        
        UlFila = Range("BE4") + 21
        AA = Range("C" & UlFila)
        
        If AA = "Reisel Sanchez" Then
            USUARIO = "SANCHEZREL"
        End If
        If AA = "Reisel" Then
            USUARIO = "REISEL SANCHEZ"
        End If
        If AA = "Leyci Colina" Then
            USUARIO = "COLINALP"
        End If
        If AA = "Elvira Alvarado" Then
            USUARIO = "ALVARADOEGO"
        End If
        If AA = "Lorena Gutierrez" Then
            USUARIO = "GUTIERREZLDX"
        End If
        If AA = "Laura Vargas" Then
            USUARIO = "VARGASLY"
        End If
        If AA = "Maria Rodriguez" Then
            USUARIO = "RODRIGUEZMDA"
        End If
        If AA = "Meyling De Sousa" Then
            USUARIO = "DESOUSAMF"
        End If
        
        If AA = "Reisel Sanchez" Then
            USUARIO2 = "sanchezrel"
        End If
        If AA = "Leyci Colina" Then
            USUARIO2 = "colinalp"
        End If
        If AA = "Elvira Alvarado" Then
            USUARIO2 = "alvaradoego"
        End If
        If AA = "Lorena Gutierrez" Then
            USUARIO2 = "gutierrezldx"
        End If
        If AA = "Laura Vargas" Then
            USUARIO2 = "vargasly"
        End If
        If AA = "Maria Rodriguez" Then
            USUARIO2 = "rodriguezmda"
        End If
        If AA = "Meyling De Sousa" Then
            USUARIO2 = "desousamf"
        End If
        

        'DESPROTEGE
        If CreateObject("WScript.Network").UserName = USUARIO Then
            Call Hoja81.Desproteger
            Range("BF2") = ""
            Range("BP1") = ""
            ActiveSheet.Shapes("Picture 47").Select   'BOTON VERDE
            Selection.ShapeRange.ZOrder msoSendToBack
            Hoja9.Visible = xlSheetHidden
            GoTo TTT
        End If
        If CreateObject("WScript.Network").UserName = USUARIO2 Then
            Call Hoja81.Desproteger
            Range("BF2") = ""
            Range("BP1") = ""
            ActiveSheet.Shapes("Picture 47").Select   'BOTON VERDE
            Selection.ShapeRange.ZOrder msoSendToBack
            Hoja9.Visible = xlSheetHidden
        Else
            MsgBox ("Solo el creador del documento lo puede desbloquear")
        End If
    Else
        ActiveSheet.Shapes("Picture 47").Select   'BOTON VERDE
        Selection.ShapeRange.ZOrder msoSendToBack
        
    End If
TTT:
End Sub
