VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Proteger Estimado"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3735
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    
    'PROTEGE LA HOJA
    ActiveSheet.Shapes("Picture 48").Select
    Selection.ShapeRange.ZOrder msoSendToBack
    Range("BF2") = 1
    Range("BP1") = ""
    Columns("BP").EntireColumn.Hidden = True
    Rows("5:5").EntireRow.Hidden = True
    Range("BP5") = TextBox1
    Hoja81.Protect Password:=Range("BP1"), DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=False
    Hoja9.Visible = xlSheetVeryHidden
    UserForm2.Hide
    TextBox1 = ""
        
End Sub

