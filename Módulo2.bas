Attribute VB_Name = "Módulo2"
Option Explicit

Sub Macro2()
Attribute Macro2.VB_Description = "Macro grabada el 23/10/2018 por Administrador"
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
' Macro grabada el 23/10/2018 por Administrador
'

'
    Range("BB11:BE11").Select
    Selection.Copy
    Range("BB17").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("BB17:BE17").Select
    Rows("17:17").RowHeight = 39
    Range("BB17:BC17").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("BE17").Select
    ActiveCell.FormulaR1C1 = "=R[-6]C"
    Range("BB17:BC17").Select
    ActiveCell.FormulaR1C1 = "=R[-6]C"
    Range("BB17:BE17").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    Rows("17:17").RowHeight = 32.25
End Sub

Sub pRUEBA()

    Range("BB17:BE17").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    Range("BB17:BC17").Select
    Selection.Interior.ColorIndex = 40
    ActiveCell.FormulaR1C1 = "=R[-6]C"
    Range("BE17").Select
    ActiveCell.FormulaR1C1 = "=R[-6]C"
    Range("BB17:BC17").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("BB14").Select

End Sub

Sub Macro3()
Attribute Macro3.VB_Description = "Macro grabada el 23/10/2018 por Administrador"
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
' Macro grabada el 23/10/2018 por Administrador
'

'
    Range("BB17:BE17").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    Range("BB17:BC17").Select
    Selection.Interior.ColorIndex = 40
    ActiveCell.FormulaR1C1 = "=R[-6]C"
    Range("BE17").Select
    ActiveCell.FormulaR1C1 = "=R[-6]C"
    Range("BB17:BC17").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("BB14").Select
End Sub
Sub Macro4()
Attribute Macro4.VB_Description = "Macro grabada el 23/10/2018 por Administrador"
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"

MsgBox CreateObject("WScript.Network").UserName

End Sub
