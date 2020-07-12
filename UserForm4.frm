VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "Datos Generales"
   ClientHeight    =   10005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11955
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox1_Click()
    CheckBox1 = True
    If CheckBox2 = True Then
        CheckBox2 = False
    End If
End Sub

Private Sub CheckBox2_Click()
    CheckBox2 = True
    If CheckBox1 = True Then
        CheckBox1 = False
    End If
End Sub

Private Sub CommandButton1_Click()
    On Error Resume Next
    Dim FILAS As Double
    'Valida llenado de los datos en Form4
    If TextBox3 = "" Then 'Nombre
        MsgBox "Debe indicar el titulo del proceso", vbCritical, "Advertencia"
        GoTo dfr
    End If
    If TextBox5 = "" Then 'Cant Renglones
        MsgBox "Debe ingresar la cantidad de renglones del proceso para continuar", vbCritical, "Advertencia"
        GoTo dfr
    End If
    If TextBox5 < 1 Then 'Cant Renglones
        MsgBox "Debe ingresar la cantidad de renglones del proceso para continuar", vbCritical, "Advertencia"
        GoTo dfr
    End If
    If TextBox6 = "" Then 'Fecha de Elaboración
        MsgBox "Debe completar los datos para continuar", vbCritical, "Advertencia"
        GoTo dfr
    End If
    If CheckBox1 = False And CheckBox2 = False Then
        MsgBox "Debe indicar el tiempo de validez a tomar (Menor/Mayor)", vbCritical, "Advertencia"
        GoTo dfr
    End If
    If CheckBox1 = True And TextBox8 = "" Then 'Validez menor
        MsgBox "Debe indicar los dias de validez para un proceso menor", vbCritical, "Advertencia"
        GoTo dfr
    End If
    If CheckBox2 = True And TextBox12 = "" Then 'Validez mayor
        MsgBox "Debe indicar los dias de validez para un proceso mayor", vbCritical, "Advertencia"
        GoTo dfr
    End If
    If TextBox11 = "" Then 'Tiempo de entrega
        MsgBox "Debe indicar el tiempo de entrega", vbCritical, "Advertencia"
        GoTo dfr
    End If
    
    If TextBox9 = "" Then 'Solicitante
        MsgBox "Debe completar los datos para continuar", vbCritical, "Advertencia"
        GoTo dfr
    End If
    If OptionButton6 = False And OptionButton8 = False And OptionButton9 = False Then 'Criterios de Numeración del Estimado
        MsgBox "Indicar Tipo de Estimado (Nacional, Para Exterior o Mixto)", vbCritical, "Advertencia"
    GoTo dfr
    End If


    Hoja81.Range("C22") = "" 'Coloca los nombres del creador del Estimado
        
        If CreateObject("WScript.Network").UserName = "SANCHEZREL" Then
            Hoja81.Range("C22") = "Reisel Sanchez"
        End If
        If CreateObject("WScript.Network").UserName = "COLINALP" Then
            Hoja81.Range("C22") = "Leyci Colina"
        End If
        If CreateObject("WScript.Network").UserName = "ALVARADOEGO" Then
            Hoja81.Range("C22") = "Elvira Alvarado"
        End If
        If CreateObject("WScript.Network").UserName = "GUTIERREZLDX" Then
            Hoja81.Range("C22") = "Lorena Gutierrez"
        End If
        If CreateObject("WScript.Network").UserName = "VARGASLY" Then
            Hoja81.Range("C22") = "Laura Vargas"
        End If
        If CreateObject("WScript.Network").UserName = "REISEL SANCHEZ" Then
            Hoja81.Range("C22") = "Reisel"
        End If
        If CreateObject("WScript.Network").UserName = "RODRIGUEZMDA" Then
            Hoja81.Range("C22") = "Maria Rodriguez"
        End If
        If CreateObject("WScript.Network").UserName = "PETITRC" Then
            Hoja81.Range("C22") = "Roquelba Petit"
        End If
        If CreateObject("WScript.Network").UserName = "sanchezrel" Then
            Hoja81.Range("C22") = "Reisel Sanchez"
        End If
        If CreateObject("WScript.Network").UserName = "colinalp" Then
            Hoja81.Range("C22") = "Leyci Colina"
        End If
        If CreateObject("WScript.Network").UserName = "alvaradoego" Then
            Hoja81.Range("C22") = "Elvira Alvarado"
        End If
        If CreateObject("WScript.Network").UserName = "gutierrezldx" Then
            Hoja81.Range("C22") = "Lorena Gutierrez"
        End If
        If CreateObject("WScript.Network").UserName = "vargasly" Then
            Hoja81.Range("C22") = "Laura Vargas"
        End If
        If CreateObject("WScript.Network").UserName = "REISEL SANCHEZ" Then
            Hoja81.Range("C22") = "Reisel"
        End If
        If CreateObject("WScript.Network").UserName = "rodriguezmda" Then
            Hoja81.Range("C22") = "Maria Rodriguez"
        End If
        If CreateObject("WScript.Network").UserName = "petitrc" Then
            Hoja81.Range("C22") = "Roquelba Petit"
        End If
        If CreateObject("WScript.Network").UserName = "desousamf" Then
            Hoja81.Range("C22") = "Meyling De Sousa"
        End If
        If CreateObject("WScript.Network").UserName = "DESOUSAMF" Then
            Hoja81.Range("C22") = "Meyling De Sousa"
        End If
        If CreateObject("WScript.Network").UserName = "navaam" Then
            Hoja81.Range("C22") = "Anais Navas"
        End If

    'Varibles de Calculos
        'Tasa Cambiaria
  '  If OptionButton1 = True Then 'No Aplicar
    '   Conversión Tasa a $
   '     Hoja81.Range("U16").FormulaLocal = "=SI(Y(M16="""";N16="""");0;SI(Y(M16<>"""";N16="""");""F.MONEDA"";SI(N16=""USD"";1;""COLOCAR"")))"
    '   Monto en $
'        Hoja81.Range("V16").FormulaLocal = "=SI(ESERROR(U16*M16);0;U16*M16)"
    '   Tasa DIPRO Momento Pedido
'        Hoja81.Range("W16") = 0
    '   Tasa SIDAMI / DICOM Momento Pedido
'        Hoja81.Range("X16") = 0
    '   Tasa DIPRO a la fecha actual
'        Hoja81.Range("Y16").FormulaLocal = "=SI(M16="""";0;BUSCARV($BE$2;'INDICE INPC'!Y:Z;2;FALSO))"
    '   Tasa DICOM a la fecha actual
'        Hoja81.Range("Z16") = 0
    '   Incremento Bs./$ Paridad
'        Hoja81.Range("AA16") = 0
    '   Costo Produccion al Momento Compra ($)
'        Hoja81.Range("AB16") = 0
    '    Aumento por Paridad del Costo (Bs.)
'        Hoja81.Range("AC16") = 0
    '   INPC del pedido
'        Hoja81.Range("AD16").FormulaLocal = "=SI(Y(L16="""";M16<>"""");""-"";SI(K16="""";""-"";BUSCARV(K16;'INDICE INPC'!D:E;2;FALSO)))"
    '   INPC Actual
'        Hoja81.Range("AE16").FormulaLocal = "=SI(Y(L16="""";M16<>"""");""-"";SI(K16="""";""-"";BUSCARV($BE$2;'INDICE INPC'!D:E;2;FALSO)))"
    '   % de Inflación Nacional Estimada
'        Hoja81.Range("AF16").FormulaLocal = "=SI(Y(L16="""";M16<>"""");""-"";SI(K16="""";""-"";((AE16-AD16)/AD16)))"
    '   CPI-U Fecha del pedido
'        Hoja81.Range("AG16").FormulaLocal = "=SI(O(M16="""";N16="""");0;SI(K16="""";""-"";((BUSCARV(K16;'INDICE INPC'!K:L;2;FALSO)))))"
    '   CPI-U Fecha actual
'        Hoja81.Range("AH16").FormulaLocal = "=SI(O(M16="""";N16="""");0;SI(K16="""";""-"";((BUSCARV($BE$2;'INDICE INPC'!K:L;2;FALSO)))))"
    '   % De inflación
'        Hoja81.Range("AI16").FormulaLocal = "=SI(O(M16="""";N16="""");0;(AH16-AG16)/AG16)"
    '   Monto $ Aumento po Inflación
'        Hoja81.Range("AJ16").FormulaLocal = "=SI(ESERROR(V16+(V16*AI16));""-"";V16*(1+AI16))"
    '    Monto en Bs
'        Hoja81.Range("AK16").FormulaLocal = "=SI(O(M16="""";N16="""");0;Y16*AJ16)"
    '   % Gastos por Nacionalización
 '       Hoja81.Range("AL16").FormulaLocal = "=SI(O(M16="""";N16="""";Y(L16<>"""";M16<>""""));0;0,288)"
    '    Precio Estimado en Bs./UNID  sin IVA / (comp Bs. + comp $)
 '       Hoja81.Range("BB16").FormulaLocal = "=SI(O(Y(L16<>"""";M16<>"""");Y(L16<>"""";M16=""""));AK16+BA16;SI(Y(L16="""";M16<>"""");BA16;0))"
    '   Costo de producción
 '       Hoja81.Range("AW16").FormulaLocal = "=SI(O(Y(L16<>"""";M16<>"""");Y(L16<>"""";M16=""""));REDONDEAR(SI($AZ$14=0,2;(L16/1,349999)*(1+AF16)+AC16;SI($AZ$14=0,3;(L16/1,4625)*(1+AF16)+AC16;0));2);0)"
    '   Gastos Administrativos
 '       Hoja81.Range("AY16").FormulaLocal = "=SI(AW16=""-"";""-"";REDONDEAR($AY$14*AW16;2))"
    '   Utilidad
'        Hoja81.Range("AZ16").FormulaLocal = "=SI(AW16=""-"";""-"";REDONDEAR((AW16+AY16)*$AZ$14;2))"
    '   Componente Estimado
'        Hoja81.Range("BA16").FormulaLocal = "=SI(K16<>"""";REDONDEAR(SI(L16<>"""";AW16+AX16+AY16+AZ16;SI(Y(L16="""";M16<>"""");AK16*(1+AL16);0));2);0)"
    '   Total Bs sin IVA
'        Hoja81.Range("BC16").FormulaLocal = "=BA16*F16"
    '   Total Componente Estimado $
'        Hoja81.Range("BD16").FormulaLocal = "=SI(O(L16="""";V16=0);0;AJ16*F16)"
    '   Total Estimado en Bs + $
'        Hoja81.Range("BE16").FormulaLocal = "=SI(K16<>"""";SI(BB16<>0;BB16*F16;BA16*F16);0)"
    '   TOTAL en bolívares (Bs.)(comp Bs. con IVA + comp $):
'        Hoja81.Range("BE11").FormulaLocal = "=SI(SUMA(BD15:BD17)=0;BE8;BE8+(BE9*(BUSCARV($BE$2;'INDICE INPC'!Y:Z;2;FALSO))))"
    '   Total en Bs para la hoja de Control Mensual
'        Hoja81.Range("BK1").FormulaLocal = "=BE11"
    
 '       End If
        
'    If OptionButton2 = True Then 'DIPRO Nacional / Exterior
    '   Conversión Tasa a $
'        Hoja81.Range("U16").FormulaLocal = "=SI(Y(M16="""";N16="""");0;SI(Y(M16<>"""";N16="""");""F.MONEDA"";SI(N16=""USD"";1;""COLOCAR"")))"
    '   Monto en $
'        Hoja81.Range("V16").FormulaLocal = "=SI(ESERROR(U16*M16);0;U16*M16)"
    '   Tasa DIPRO Momento Pedido
'        Hoja81.Range("W16").FormulaLocal = "=SI(K16="""";0;BUSCARV(K16;'INDICE INPC'!Y:Z;2;FALSO))"
    '   Tasa SIDAMI / DICOM Momento Pedido
'        Hoja81.Range("X16") = 0
    '   Tasa DIPRO a la fecha actual
'        Hoja81.Range("Y16").FormulaLocal = "=SI(K16="""";0;BUSCARV($BE$2;'INDICE INPC'!Y:Z;2;FALSO))"
    '   Tasa DICOM a la fecha actual
'        Hoja81.Range("Z16") = 0
    '    Incremento Bs./$ Paridad
        'Hoja81.Range("AA16").FormulaLocal = "=SI(L16<>"""";Y16-W16;0)"
    '    Costo Produccion al Momento Compra ($)
     '   Hoja81.Range("AB16").FormulaLocal = "=SI(L16="""";0;SI($AZ$14=0,2;(L16/1,349999)/W16;(L16/1,4625)/W16))"
    '    Aumento por Paridad del Costo (Bs.)
        'Hoja81.Range("AC16").FormulaLocal = "=AB16*AA16"
    '   INPC del pedido
'        Hoja81.Range("AD16").FormulaLocal = "=SI(Y(L16="""";M16<>"""");""-"";SI(K16="""";""-"";BUSCARV(K16;'INDICE INPC'!D:E;2;FALSO)))"
    '   INPC Actual
'        Hoja81.Range("AE16").FormulaLocal = "=SI(Y(L16="""";M16<>"""");""-"";SI(K16="""";""-"";BUSCARV($BE$2;'INDICE INPC'!D:E;2;FALSO)))"
   '   % de Inflación Nacional Estimada
'        Hoja81.Range("AF16").FormulaLocal = "=SI(Y(L16="""";M16<>"""");""-"";SI(K16="""";""-"";((AE16-AD16)/AD16)))"
    '   CPI-U Fecha del pedido
'        Hoja81.Range("AG16").FormulaLocal = "=SI(O(M16="""";N16="""");0;SI(K16="""";""-"";((BUSCARV(K16;'INDICE INPC'!K:L;2;FALSO)))))"
    '   CPI-U Fecha actual
'        Hoja81.Range("AH16").FormulaLocal = "=SI(O(M16="""";N16="""");0;SI(K16="""";""-"";((BUSCARV($BE$2;'INDICE INPC'!K:L;2;FALSO)))))"
    '   % De inflación
'        Hoja81.Range("AI16").FormulaLocal = "=SI(O(M16="""";N16="""");0;(AH16-AG16)/AG16)"
    '   Monto $ Aumento po Inflación
'        Hoja81.Range("AJ16").FormulaLocal = "=SI(ESERROR(V16+(V16*AI16));""-"";V16*(1+AI16))"
   '     Monto en Bs
'        Hoja81.Range("AK16").FormulaLocal = "=SI(O(M16="""";N16="""");0;Y16*AJ16)"
    '   % Gastos por Nacionalización
'        Hoja81.Range("AL16").FormulaLocal = "=SI(O(M16="""";N16="""";Y(L16<>"""";M16<>""""));0;0,288)"
    '    Precio Estimado en Bs./UNID  sin IVA / (comp Bs. + comp $)
'        Hoja81.Range("BB16").FormulaLocal = "=SI(O(L16<>"""";M16<>"""");SI(Y(L16<>"""";M16="""");BA16;SI(Y(L16<>"""";N16<>"""");BA16+AK16;SI(Y(L16="""";N16<>"""");BA16)));0)"
    '   Costo de producción
'        Hoja81.Range("AW16").FormulaLocal = "=SI(O(Y(L16<>"""";M16<>"""");Y(L16<>"""";M16=""""));REDONDEAR(SI($AZ$14=0,2;(L16/1,349999)*(1+AF16)+AC16;SI($AZ$14=0,3;(L16/1,4625)*(1+AF16)+AC16;0));2);0)"
    '   Gastos Administrativos
'        Hoja81.Range("AY16").FormulaLocal = "=SI(AW16=""-"";""-"";REDONDEAR($AY$14*AW16;2))"
    '   Utilidad
'        Hoja81.Range("AZ16").FormulaLocal = "=SI(AW16=""-"";""-"";REDONDEAR((AW16+AY16)*$AZ$14;2))"
    '   Componente Estimado
'        Hoja81.Range("BA16").FormulaLocal = "=SI(K16<>"""";REDONDEAR(SI(L16<>"""";AW16+AX16+AY16+AZ16;SI(Y(L16="""";M16<>"""");AK16*(1+AL16);0));2);0)"
    '   Total Bs sin IVA
'        Hoja81.Range("BC16").FormulaLocal = "=BA16*F16"
    '   Total Componente Estimado $
'        Hoja81.Range("BD16").FormulaLocal = "=SI(O(L16="""";V16=0);0;AJ16*F16)"
    '   Total Estimado en Bs + $
'        Hoja81.Range("BE16").FormulaLocal = "=SI(K16<>"""";SI(BB16<>0;BB16*F16;BA16*F16);0)"
    '   TOTAL en bolívares (Bs.)(comp Bs. con IVA + comp $):
'        Hoja81.Range("BE11").FormulaLocal = "=SI(SUMA(BD15:BD17)=0;BE8;BE8+(BE9*(BUSCARV($BE$2;'INDICE INPC'!Y:Z;2;FALSO))))"
    '   Total en Bs para la hoja de Control Mensual
'        Hoja81.Range("BK1").FormulaLocal = "=BE11"

'    End If
    
    If OptionButton6 = True Then 'PROCESO NACIONAL
    '   Conversión Tasa a $
        Hoja81.Range("U16").FormulaLocal = "=SI(Y(M16="""";N16="""");0;SI(Y(M16<>"""";N16="""");""F.MONEDA"";SI(N16=""USD"";1;""COLOCAR"")))"
    '   Monto en $
        Hoja81.Range("V16").FormulaLocal = "=SI(ESERROR(U16*M16);0;U16*M16)"
    '   Tasa DIPRO Momento Pedido
 '       Hoja81.Range("W16").FormulaLocal = "=SI(O(K16="""";L16="""");0;BUSCARV(K16;'INDICE INPC'!Y:Z;2;FALSO))"
    '   Tasa SIDAMI / DICOM Momento Pedido
        Hoja81.Range("X16").FormulaLocal = "=SI(O(K16="""";M16="""");"""";BUSCARV(K16;'INDICE INPC'!R:S;2;FALSO))"
    '   Tasa DIPRO a la fecha actual
'        Hoja81.Range("Y16").FormulaLocal = "=SI(O(K16="""";L16="""");0;BUSCARV($BE$2;'INDICE INPC'!Y:Z;2;FALSO))"
    '   Tasa DICOM a la fecha actual
        Hoja81.Range("Z16").FormulaLocal = "=SI(O(K16="""";M16="""");"""";BUSCARV($BE$2;'INDICE INPC'!R:S;2;FALSO))"
   '    Incremento Bs./$ Paridad
        'Hoja81.Range("AA16").FormulaLocal = "=SI(L16<>"""";Y16-W16;0)"
    '    Costo Produccion al Momento Compra ($)
      '  Hoja81.Range("AB16").FormulaLocal = "=SI(O(L16="""";Y16=0);0;SI($AZ$14=0,2;(L16/1,349999)/W16;(L16/1,4625)/W16))"
    '    Aumento por Paridad del Costo (Bs.)
       ' Hoja81.Range("AC16").FormulaLocal = "=AB16*AA16"
    '   INPC del pedido
        Hoja81.Range("AD16").FormulaLocal = "=SI(Y(L16="""";M16<>"""");""-"";SI(K16="""";""-"";BUSCARV(K16;'INDICE INPC'!D:E;2;FALSO)))"
    '   INPC Actual
        Hoja81.Range("AE16").FormulaLocal = "=SI(Y(L16="""";M16<>"""");""-"";SI(K16="""";""-"";BUSCARV($BE$2;'INDICE INPC'!D:E;2;FALSO)))"
   '   % de Inflación Nacional Estimada
        Hoja81.Range("AF16").FormulaLocal = "=SI(Y(L16="""";M16<>"""");""-"";SI(K16="""";""-"";((AE16-AD16)/AD16)))"
    '   CPI-U Fecha del pedido
        Hoja81.Range("AG16").FormulaLocal = "=SI(O(M16="""";N16="""");0;SI(K16="""";""-"";((BUSCARV(K16;'INDICE INPC'!K:L;2;FALSO)))))"
    '   CPI-U Fecha actual
        Hoja81.Range("AH16").FormulaLocal = "=SI(O(M16="""";N16="""");0;SI(K16="""";""-"";((BUSCARV($BE$2;'INDICE INPC'!K:L;2;FALSO)))))"
    '   % De inflación
        Hoja81.Range("AI16").FormulaLocal = "=SI(O(M16="""";N16="""");0;(AH16-AG16)/AG16)"
    '   Monto $ Aumento po Inflación
        Hoja81.Range("AJ16").FormulaLocal = "=SI(ESERROR(V16+(V16*AI16));""-"";V16*(1+AI16))"
    '    Monto en Bs
        Hoja81.Range("AK16").FormulaLocal = "=SI(O(M16="""";N16="""");0;SI(Y(M16<>"""";L16<>"""");Z16*AJ16;SI(L16="""";AJ16*Z16)))"
    '   % Gastos por Nacionalización
        Hoja81.Range("AL16").FormulaLocal = "=SI(O(M16="""";N16="""";Y(L16<>"""";M16<>""""));0;0,288)"
    '    Precio Estimado en Bs./UNID  sin IVA / (comp Bs. + comp $)
        Hoja81.Range("BB16").FormulaLocal = "=SI(O(L16<>"""";M16<>"""");SI(Y(L16<>"""";M16="""");BA16;SI(Y(L16<>"""";N16<>"""");BA16+AK16;SI(Y(L16="""";N16<>"""");BA16)));0)"
    '   Costo de producción
        Hoja81.Range("AW16").FormulaLocal = "=SI(O(Y(L16<>"""";M16<>"""");Y(L16<>"""";M16=""""));REDONDEAR(SI($AZ$14=0,2;(L16/1,349999)*(1+AF16)+AC16;SI($AZ$14=0,3;(L16/1,4625)*(1+AF16)+AC16;0));2);0)"
    '   Gastos Administrativos
        Hoja81.Range("AY16").FormulaLocal = "=SI(AW16=""-"";""-"";REDONDEAR($AY$14*AW16;2))"
    '   Utilidad
        Hoja81.Range("AZ16").FormulaLocal = "=SI(AW16=""-"";""-"";REDONDEAR((AW16+AY16)*$AZ$14;2))"
    '   Componente Estimado
        Hoja81.Range("BA16").FormulaLocal = "=SI(K16<>"""";REDONDEAR(SI(L16<>"""";AW16+AX16+AY16+AZ16;SI(Y(L16="""";M16<>"""");AK16*(1+AL16);0));2);0)"
    '   Total Bs sin IVA
        Hoja81.Range("BC16").FormulaLocal = "=BA16*F16"
    '   Total Componente Estimado $
        Hoja81.Range("BD16").FormulaLocal = "=SI(O(L16="""";V16=0);0;AJ16*F16)"
    '   Total Estimado en Bs + $
        Hoja81.Range("BE16").FormulaLocal = "=SI(K16<>"""";SI(BB16<>0;BB16*F16;BA16*F16);0)"
    '   Total Estimado en Bs.S
        'Hoja81.Range("BF16").FormulaLocal = "=BE16/100000"
    '   TOTAL en bolívares (Bs.)(comp Bs. con IVA + comp $):
        Hoja81.Range("BE11").FormulaLocal = "=SI(SUMA(BD15:BD17)=0;BE8;BE8+(BE9*(BUSCARV($BE$2;'INDICE INPC'!R:S;2;FALSO))))"
    '   TOTAL en bolívares (Bs.)(comp Bs. con IVA + comp $):
        'Hoja81.Range("BF11").FormulaLocal = "=BE11/100000"
    '   Total en Bs para la hoja de Control Mensual
        Hoja81.Range("BK1").FormulaLocal = "=BE11"
        
    '   TOTAL AL FINAL
    Worksheets("Estimado Base").Rows("17:17").Hidden = True
    Worksheets("Estimado Base").Rows("18:18").Hidden = False
            Range("BB18:BE18").Select
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
    Range("BB18:BC18").Select
    Selection.Interior.ColorIndex = 40
    ActiveCell.FormulaR1C1 = "=R[-7]C"
    Range("BE18").Select
    ActiveCell.FormulaR1C1 = "=R[-7]C"
    Range("BB18:BC18").Select
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
    
    Rows("18:18").RowHeight = 30
    Rows("20:20").RowHeight = 54
        
        
    End If
'    If OptionButton5 = True Then 'DIPRO Nacional / DICON Nacional
    '   Conversión Tasa a $
'        Hoja81.Range("U16").FormulaLocal = "=SI(Y(M16="""";N16="""");0;SI(Y(M16<>"""";N16="""");""F.MONEDA"";SI(N16=""USD"";1;""COLOCAR"")))"
    '   Monto en $
'        Hoja81.Range("V16").FormulaLocal = "=SI(ESERROR(U16*M16);0;U16*M16)"
    '   Tasa DIPRO Momento Pedido
'        Hoja81.Range("W16") = 0
    '   Tasa SIDAMI / DICOM Momento Pedido
'        Hoja81.Range("X16").FormulaLocal = "=SI(K16="""";0;BUSCARV(K16;'INDICE INPC'!R:S;2;FALSO))"
    '   Tasa DIPRO a la fecha actual
'        Hoja81.Range("Y16") = 0
    '   Tasa DICOM a la fecha actual
'        Hoja81.Range("Z16").FormulaLocal = "=SI(K16="""";0;BUSCARV($BE$2;'INDICE INPC'!R:S;2;FALSO))"
    '    Incremento Bs./$ Paridad
        'Hoja81.Range("AA16").FormulaLocal = "=SI(L16<>"""";Z16-X16;0)"
    '    Costo Produccion al Momento Compra ($)
      '  Hoja81.Range("AB16").FormulaLocal = "=SI(L16="""";0;SI($AZ$14=0,2;(L16/1,349999)/X16;(L16/1,4625)/X16))"
    '    Aumento por Paridad del Costo (Bs.)
        'Hoja81.Range("AC16").FormulaLocal = "=AB16*AA16"
    '   INPC del pedido
'        Hoja81.Range("AD16").FormulaLocal = "=SI(Y(L16="""";M16<>"""");""-"";SI(K16="""";""-"";BUSCARV(K16;'INDICE INPC'!D:E;2;FALSO)))"
    '   INPC Actual
'        Hoja81.Range("AE16").FormulaLocal = "=SI(Y(L16="""";M16<>"""");""-"";SI(K16="""";""-"";BUSCARV($BE$2;'INDICE INPC'!D:E;2;FALSO)))"
    '   % de Inflación Nacional Estimada
'        Hoja81.Range("AF16").FormulaLocal = "=SI(Y(L16="""";M16<>"""");""-"";SI(K16="""";""-"";((AE16-AD16)/AD16)))"
    '   CPI-U Fecha del pedido
'        Hoja81.Range("AG16").FormulaLocal = "=SI(O(M16="""";N16="""");0;SI(K16="""";""-"";((BUSCARV(K16;'INDICE INPC'!K:L;2;FALSO)))))"
    '   CPI-U Fecha actual
'        Hoja81.Range("AH16").FormulaLocal = "=SI(O(M16="""";N16="""");0;SI(K16="""";""-"";((BUSCARV($BE$2;'INDICE INPC'!K:L;2;FALSO)))))"
    '   % De inflación
'        Hoja81.Range("AI16").FormulaLocal = "=SI(O(M16="""";N16="""");0;(AH16-AG16)/AG16)"
    '   Monto $ Aumento po Inflación
'        Hoja81.Range("AJ16").FormulaLocal = "=SI(ESERROR(V16+(V16*AI16));""-"";V16*(1+AI16))"
    '    Monto en Bs
'        Hoja81.Range("AK16").FormulaLocal = "=SI(O(M16="""";N16="""");0;Z16*AJ16)"
    '   % Gastos por Nacionalización
'        Hoja81.Range("AL16").FormulaLocal = "=SI(O(M16="""";N16="""";Y(L16<>"""";M16<>""""));0;0,288)"
    '   Precio Estimado en Bs./UNID  sin IVA / (comp Bs. + comp $)
'        Hoja81.Range("BB16").FormulaLocal = "=SI(O(L16<>"""";M16<>"""");SI(Y(L16<>"""";M16="""");BA16;SI(Y(L16<>"""";N16<>"""");BA16+AK16;SI(Y(L16="""";N16<>"""");BA16)));0)"
    '   Costo de producción
'        Hoja81.Range("AW16").FormulaLocal = "=SI(O(Y(L16<>"""";M16<>"""");Y(L16<>"""";M16=""""));REDONDEAR(SI($AZ$14=0,2;(L16/1,349999)*(1+AF16)+AC16;SI($AZ$14=0,3;(L16/1,4625)*(1+AF16)+AC16;0));2);0)"
    '   Gastos Administrativos
'        Hoja81.Range("AY16").FormulaLocal = "=SI(AW16=""-"";""-"";REDONDEAR($AY$14*AW16;2))"
    '   Utilidad
'        Hoja81.Range("AZ16").FormulaLocal = "=SI(AW16=""-"";""-"";REDONDEAR((AW16+AY16)*$AZ$14;2))"
    '   Componente Estimado
'        Hoja81.Range("BA16").FormulaLocal = "=SI(K16<>"""";REDONDEAR(SI(L16<>"""";AW16+AX16+AY16+AZ16;SI(Y(L16="""";M16<>"""");AK16*(1+AL16);0));2);0)"
    '   Total Bs sin IVA
'        Hoja81.Range("BC16").FormulaLocal = "=BA16*F16"
    '   Total Componente Estimado $
'        Hoja81.Range("BD16").FormulaLocal = "=SI(O(L16="""";V16=0);0;AJ16*F16)"
    '   Total Estimado en Bs + $
'        Hoja81.Range("BE16").FormulaLocal = "=SI(K16<>"""";SI(BB16<>0;BB16*F16;BA16*F16);0)"
    '   TOTAL en bolívares (Bs.)(comp Bs. con IVA + comp $):
'        Hoja81.Range("BE11").FormulaLocal = "=SI(SUMA(BD15:BD17)=0;BE8;BE8+(BE9*(BUSCARV($BE$2;'INDICE INPC'!R:S;2;FALSO))))"
    '   Total en Bs para la hoja de Control Mensual
'        Hoja81.Range("BK1").FormulaLocal = "=BE11"
    
'        End If
           
           
         ' Incremento Salarial
 '   If OptionButton3 = True Then 'No Aplica
 '       Hoja81.Range("AM16") = 0
 '       Hoja81.Range("AN16") = 0
 '       Hoja81.Range("AO16") = 0
 '       Hoja81.Range("AX16") = 0
 '   End If
 '   If OptionButton4 = True Then 'Aplica
'        Hoja81.Range("AM16").FormulaLocal = "=SI(L16="""";0;BUSCARV(K16;'INDICE INPC'!AE:AF;2;FALSO))"
'        Hoja81.Range("AN16").FormulaLocal = "=SI(L16="""";0;BUSCARV($BE$2;'INDICE INPC'!AE:AF;2;FALSO))"
'        Hoja81.Range("AO16").FormulaLocal = "=SI(L16="""";0;(AN16-AM16)/AM16)"
'        Hoja81.Range("AX16").FormulaLocal = "=SI(O(K16="""";AO16=0);0;(((SI($AZ$14=0,2;(L16/1,349999);(L16/1,4625)))*12,5%)*(1+AO16)))"
'    End If
       
        ' Analisis Salario
'    If OptionButton7 = True Then
    '   Conversión Tasa a $
'        Hoja81.Range("U16").FormulaLocal = "=SI(Y(M16="""";N16="""");0;SI(Y(M16<>"""";N16="""");""F.MONEDA"";SI(N16=""USD"";1;""COLOCAR"")))"
    '   Monto en $
'        Hoja81.Range("V16").FormulaLocal = "=SI(ESERROR(U16*M16);0;U16*M16)"
    '   Tasa DIPRO Momento Pedido
'        Hoja81.Range("W16") = 0
    '   Tasa SIDAMI / DICOM Momento Pedido
'        Hoja81.Range("X16") = 0
    '   Tasa DIPRO a la fecha actual
'        Hoja81.Range("Y16") = 0
    '   Tasa DICOM a la fecha actual
'        Hoja81.Range("Z16") = 0
    '    Incremento Bs./$ Paridad
'        Hoja81.Range("AA16") = 0
    '    Costo Produccion al Momento Compra ($)
'        Hoja81.Range("AB16") = 0
    '    Aumento por Paridad del Costo (Bs.)
 '       Hoja81.Range("AC16") = 0
    '   % de Inflación Nacional Estimada
 '       Hoja81.Range("AF16") = 0
    '   CPI-U Fecha del pedido
'        Hoja81.Range("AG16").FormulaLocal = "=SI(O(M16=0;N16=0);0;SI(K16="""";""-"";((BUSCARV(K16;'INDICE INPC'!K:L;2;FALSO)))))"
    '   CPI-U Fecha actual
'        Hoja81.Range("AH16").FormulaLocal = "=SI(O(M16=0;N16=0);0;SI(K16="""";""-"";((BUSCARV($BE$2;'INDICE INPC'!K:L;2;FALSO)))))"
    '    Monto en Bs
'        Hoja81.Range("AK16") = 0
    '   % Gastos por Nacionalización
'        Hoja81.Range("AL16").FormulaLocal = "=SI(O(M16="""";N16="""";Y(L16<>"""";M16<>""""));0;0,288)"
    '    Precio Estimado en Bs./UNID  sin IVA / (comp Bs. + comp $)
'        Hoja81.Range("BB16").FormulaLocal = "=SI(O(L16<>"""";M16<>"""");SI(Y(L16<>"""";M16="""");BA16;SI(Y(L16<>"""";N16<>"""");BA16+(AJ16*Z16);SI(Y(L16="""";N16<>"""");BA16)));0)"
        'Salario Minimo Pedido
'        Hoja81.Range("AM16").FormulaLocal = "=SI(L16="""";0;BUSCARV(K16;'INDICE INPC'!AE:AF;2;FALSO))"
        'Salario Mínimo Actual
'        Hoja81.Range("AN16").FormulaLocal = "=SI(L16="""";0;BUSCARV($BE$2;'INDICE INPC'!AE:AF;2;FALSO))"
        '% de Incremento Salarial
'        Hoja81.Range("AO16").FormulaLocal = "=SI(L16="""";0;(AN16-AM16)/AM16)"
        'Costo Por Incremento Salarial. en Bs./UNID Estimado
'        Hoja81.Range("AX16").FormulaLocal = "=SI(O(K16="""";AO16=0);0;(((SI($AZ$14=0,2;(L16/1,349999);(L16/1,4625)))*12,5%)*(1+AO16)))"
    '   Costo de producción
'        Hoja81.Range("AW16").FormulaLocal = "=SI(O(Y(L16<>"""";M16<>"""");Y(L16<>"""";M16=""""));REDONDEAR(SI($AZ$14=0,2;(L16/1,349999)*(1+AF16)+AC16;SI($AZ$14=0,3;(L16/1,4625)*(1+AF16)+AC16;0));2);0)"
    '   Gastos Administrativos
'        Hoja81.Range("AY16").FormulaLocal = "=SI(AW16=""-"";""-"";REDONDEAR($AY$14*AW16;2))"
    '   Utilidad
'        Hoja81.Range("AZ16").FormulaLocal = "=SI(AW16=""-"";""-"";REDONDEAR((AW16+AY16)*$AZ$14;2))"
    '   Componente Estimado
 '       Hoja81.Range("BA16").FormulaLocal = "=SI(K16<>"""";REDONDEAR(SI(L16<>"""";AW16+AX16+AY16+AZ16;SI(Y(L16="""";M16<>"""");AK16*(1+AL16);0));2);0)"
    '   Total Bs sin IVA
 '       Hoja81.Range("BC16").FormulaLocal = "=BA16*F16"
    '   Total Componente Estimado $
 '       Hoja81.Range("BD16").FormulaLocal = "=SI(Y(L16="""";V16=0);0;AJ16*F16)"
    '   Total Estimado en Bs + $
  '      Hoja81.Range("BE16").FormulaLocal = "=SI(K16<>"""";SI(BB16<>0;BB16*F16;BA16*F16);0)"
        
        'Ocultar columnas
'        Columns("O:AB").EntireColumn.Hidden = True
'        Columns("AD:AE").EntireColumn.Hidden = True
'        Columns("AG:AL").EntireColumn.Hidden = True
 '       Columns("AP:AV").EntireColumn.Hidden = True
 '       Columns("BB:BB").EntireColumn.Hidden = True
 '       Columns("BD:BE").EntireColumn.Hidden = True
  '      Columns("B:C").EntireColumn.Hidden = True
 '       Columns("E:E").EntireColumn.Hidden = True
 '       Columns("H:I").EntireColumn.Hidden = True
 '       Columns("M:N").EntireColumn.Hidden = True
        
'    End If
        
        
    If OptionButton9 = True Then 'ESTIMADO PROCESO MIXTO DIPRO Nacional / DICON Exterior
    '   Conversión Tasa a $
        Hoja81.Range("U16").FormulaLocal = "=SI(Y(M16="""";N16="""");0;SI(Y(M16<>"""";N16="""");""F.MONEDA"";SI(N16=""USD"";1;""COLOCAR"")))"
    '   Monto en $
        Hoja81.Range("V16").FormulaLocal = "=SI(ESERROR(U16*M16);0;U16*M16)"
    '   Tasa DIPRO Momento Pedido
   '     Hoja81.Range("W16").FormulaLocal = "=SI(O(K16="""";L16="""");0;BUSCARV(K16;'INDICE INPC'!Y:Z;2;FALSO))"
    '   Tasa SIDAMI / DICOM Momento Pedido
        Hoja81.Range("X16").FormulaLocal = "=SI(O(K16="""";M16="""");0;BUSCARV(K16;'INDICE INPC'!R:S;2;FALSO))"
    '   Tasa DIPRO a la fecha actual
      '  Hoja81.Range("Y16").FormulaLocal = "=SI(O(K16="""";L16="""");0;BUSCARV($BE$2;'INDICE INPC'!Y:Z;2;FALSO))"
    '   Tasa DICOM a la fecha actual
        Hoja81.Range("Z16").FormulaLocal = "=SI(O(K17="""";M17="""");0;BUSCARV($BE$2;'INDICE INPC'!R:S;2;FALSO))"
   '    Incremento Bs./$ Paridad
       ' Hoja81.Range("AA16").FormulaLocal = "=SI(L16<>"""";Y16-W16;0)"
    '    Costo Produccion al Momento Compra ($)
     '   Hoja81.Range("AB16").FormulaLocal = "=SI(O(L16="""";Y16=0);0;SI($AZ$14=0,2;(L16/1,349999)/W16;(L16/1,4625)/W16))"
    '    Aumento por Paridad del Costo (Bs.)
        'Hoja81.Range("AC16").FormulaLocal = "=AB16*AA16"
    '   INPC del pedido
        Hoja81.Range("AD16").FormulaLocal = "=SI(Y(L16="""";M16<>"""");""-"";SI(K16="""";""-"";BUSCARV(K16;'INDICE INPC'!D:E;2;FALSO)))"
    '   INPC Actual
        Hoja81.Range("AE16").FormulaLocal = "=SI(Y(L16="""";M16<>"""");""-"";SI(K16="""";""-"";BUSCARV($BE$2;'INDICE INPC'!D:E;2;FALSO)))"
   '   % de Inflación Nacional Estimada
        Hoja81.Range("AF16").FormulaLocal = "=SI(Y(L16="""";M16<>"""");""-"";SI(K16="""";""-"";((AE16-AD16)/AD16)))"
    '   CPI-U Fecha del pedido
        Hoja81.Range("AG16").FormulaLocal = "=SI(O(M16="""";N16="""");0;SI(K16="""";""-"";((BUSCARV(K16;'INDICE INPC'!K:L;2;FALSO)))))"
    '   CPI-U Fecha actual
        Hoja81.Range("AH16").FormulaLocal = "=SI(O(M16="""";N16="""");0;SI(K16="""";""-"";((BUSCARV($BE$2;'INDICE INPC'!K:L;2;FALSO)))))"
    '   % De inflación
        Hoja81.Range("AI16").FormulaLocal = "=SI(O(M16="""";N16="""");0;(AH16-AG16)/AG16)"
    '   Monto $ Aumento po Inflación
        Hoja81.Range("AJ16").FormulaLocal = "=SI(ESERROR(V16+(V16*AI16));""-"";V16*(1+AI16))"
    '    Monto en Bs
    '    Hoja81.Range("AK16").FormulaLocal = ""
    '   % Gastos por Nacionalización
      '  Hoja81.Range("AL16").FormulaLocal = "=SI(O(M16="""";N16="""";Y(L16<>"""";M16<>""""));0;0,288)"
    '   Costo de producción
        Hoja81.Range("AW16").FormulaLocal = "=SI(O(Y(L16<>"""";M16<>"""");Y(L16<>"""";M16=""""));REDONDEAR(SI($AZ$14=0,2;(L16/1,349999)*(1+AF16)+AC16;SI($AZ$14=0,3;(L16/1,4625)*(1+AF16)+AC16;0));2);0)"
    '   Gastos Administrativos
        Hoja81.Range("AY16").FormulaLocal = "=SI(AW16=""-"";""-"";REDONDEAR($AY$14*AW16;2))"
    '   Utilidad
        Hoja81.Range("AZ16").FormulaLocal = "=SI(AW16=""-"";""-"";REDONDEAR((AW16+AY16)*$AZ$14;2))"
    '   Componente Estimado
        Hoja81.Range("BA16").FormulaLocal = "=SI(K16<>"""";REDONDEAR(SI(L16<>"""";AW16+AX16+AY16+AZ16;SI(Y(L16="""";M16<>"""");AK16*(1+AL16);0));2);0)"
    '   Precio Estimado en Bs./UNID  sin IVA / (comp Bs. + comp $)
        Hoja81.Range("BB15").FormulaR1C1 = "Componente" & Chr(10) & "Estimado" & Chr(10) & "en $/UNID"
        Hoja81.Range("BB16").FormulaLocal = "=SI(M16="""";0;AJ16)"
        Hoja81.Range("BB16").NumberFormat = _
            "_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* ""-""??_);_(@_)"
    '   Total Bs sin IVA
        Hoja81.Range("BC16").FormulaLocal = "=BA16*F16"
    '   Ajuste de Celda
        Hoja81.Range("BD15").ClearContents
        Hoja81.Range("BD16").ClearContents
    '   Total Estimado en Bs + $
        Hoja81.Range("BE15").FormulaR1C1 = "TOTAL" & Chr(10) & "Componente" & Chr(10) & "Estimado en $"
        Hoja81.Range("BE16").FormulaLocal = "=SI(M16="""";0;AJ16*F16)"
        Hoja81.Range("BE16").NumberFormat = _
            "_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* ""-""??_);_(@_)"
    '   Componente Total Estimado en $
        Hoja81.Range("BE9").FormulaLocal = "=SUMA(BE16:BE17)"
    '   Total en Bs (Comp Bs. con IVA + comp $)
        Hoja81.Range("BE10").ClearContents
    '   TOTAL en bolívares (Bs.)(comp Bs. con IVA + comp $):
        Hoja81.Range("BE11").ClearContents
    '   TOTAL Actual en Bs.Sin IVA (Comp)
        'Hoja81.Range("BF15").FormulaR1C1 = "TOTAL Actual" & Chr(10) & "en Bs. sin IVA" & Chr(10) & "(Comp Bs)"
        'Hoja81.Range("BF16").FormulaLocal = "=BC16/100000"
    '   Total en Bs para la hoja de Control Mensual
        Hoja81.Range("BK1").FormulaLocal = "=BE8"


        Range("BB9:BE9").Select
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
    
    '   UCAU
        Hoja81.Range("BE13").ClearContents
        Hoja81.Range("BE14").ClearContents
        Range("BE13").Select
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            Selection.Borders(xlEdgeLeft).LineStyle = xlNone
            Selection.Borders(xlEdgeTop).LineStyle = xlNone
            Selection.Borders(xlEdgeBottom).LineStyle = xlNone
            Selection.Borders(xlEdgeRight).LineStyle = xlNone
            Selection.Borders(xlInsideVertical).LineStyle = xlNone
            Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
            Selection.Interior.ColorIndex = xlNone
        Range("BE14").Select
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            Selection.Borders(xlEdgeLeft).LineStyle = xlNone
            Selection.Borders(xlEdgeTop).LineStyle = xlNone
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        Selection.Borders(xlEdgeRight).LineStyle = xlNone
        
        Worksheets("Estimado Base").Rows("10:11").Hidden = True
    End If
           
        'ESTIMADO PARA EL EXTERIOR
        
    If OptionButton8 = True Then
    '   Se elimina el UCAU
        Range("BE13:BE14").Select
        Selection.ClearContents
        Selection.Interior.ColorIndex = xlNone
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        Selection.Borders(xlEdgeLeft).LineStyle = xlNone
        Selection.Borders(xlEdgeTop).LineStyle = xlNone
        With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
        End With
        Selection.Borders(xlEdgeRight).LineStyle = xlNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    '   Conversión Tasa a $
        Hoja81.Range("U16").FormulaLocal = "=SI(Y(M16="""";N16="""");0;SI(Y(M16<>"""";N16="""");""F.MONEDA"";SI(N16=""USD"";1;""COLOCAR"")))"
    '   Monto en $
        Hoja81.Range("V16").FormulaLocal = "=SI(ESERROR(U16*M16);0;U16*M16)"
    '   Tasa DIPRO Momento Pedido
    '    Hoja81.Range("W16").FormulaLocal = "=SI(K16="""";0;BUSCARV(K16;'INDICE INPC'!Y:Z;2;FALSO))"
    '   Tasa SIDAMI / DICOM Momento Pedido
        Hoja81.Range("X16").FormulaLocal = "=SI(K16="""";0;BUSCARV(K16;'INDICE INPC'!Y:Z;2;FALSO))"
    '   Tasa DIPRO a la fecha actual
  '      Hoja81.Range("Y16").FormulaLocal = "=SI(K16="""";0;BUSCARV($BE$2;'INDICE INPC'!Y:Z;2;FALSO)"
    '   Tasa DICOM a la fecha actual
        Hoja81.Range("Z16") = 10
    '    Costo Produccion al Momento Compra ($)
    '    Hoja81.Range("AB16").FormulaLocal = "=SI(Y(L16<>"""";M16<>"""");V16;SI(L16="""";SI(M16="""";0;V16);(L16/1,349999)/W16))"
    '    INPC fecha del pedido
        Hoja81.Range("AD16") = 0
    '    INPC fecha actual
        Hoja81.Range("AE16") = 0
    '   % de Inflación Nacional Estimada
        Hoja81.Range("AF16") = 0
    '   CPI-U Fecha del Pedido
        Hoja81.Range("AG16").FormulaLocal = "=SI(K16="""";""-"";((BUSCARV(K16;'INDICE INPC'!K:L;2;FALSO))))"
    '   CPI-U Fecha actual Pedido
        Hoja81.Range("AH16").FormulaLocal = "=SI(K16="""";""-"";((BUSCARV($BE$2;'INDICE INPC'!K:L;2;FALSO))))"
    '   Inflación del Exterior
        Hoja81.Range("AI16").FormulaLocal = "=SI(K16="""";0;((AH16-AG16)/AG16))"
    '   Monto $ Aumento Inflación
        Hoja81.Range("AJ16").FormulaLocal = "=(SI(Y(L16<>"""";M16<>"""");V16;SI(L16="""";SI(M16="""";0;V16);(L16/1,349999)/X16)))*(1+AI16)"
     '  Monto en Bs
        Hoja81.Range("AK16") = 0
     '  Gastos por Nacionalización
        Hoja81.Range("AL16") = 0
    '   Costo de producción
        Hoja81.Range("AW16") = 0
    '   Gastos Administrativos
        Hoja81.Range("AY16") = 0
    '   Utilidad
        Hoja81.Range("AZ16") = 0
    '   Estimado UND $
        Hoja81.Range("BA15").FormulaR1C1 = "Estimado" & Chr(10) & "$/UND"
        Hoja1.Range("N2").FormulaR1C1 = "Estimado" & Chr(10) & "$/UND"
        Hoja81.Range("BA16").FormulaLocal = "=AJ16"
        Hoja81.Range("BA16").Select
        Selection.NumberFormat = _
        "_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* ""-""??_);_(@_)"
    '   Estimado UND BS
        Hoja81.Range("BB15").FormulaR1C1 = "Estimado" & Chr(10) & "Bs./UND"
        Hoja1.Range("O2").FormulaR1C1 = "Estimado" & Chr(10) & "Bs./UND"
        Hoja81.Range("BB16").FormulaLocal = "=SI(ESERROR(AJ16*10);0;AJ16*10)"
    '   Total Estimado $
        Hoja81.Range("BC15").FormulaR1C1 = "TOTAL" & Chr(10) & "Estimado" & Chr(10) & "$"
        Hoja81.Range("BC16").FormulaLocal = "=SI(AJ16=0;0;BA16*F16)"
        Hoja81.Range("BC16").Select
        Selection.NumberFormat = _
        "_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* ""-""??_);_(@_)"
        Selection.Font.Bold = True
        
    '   Componente $
        Hoja81.Range("BD16") = 0
    '   Total Estimado en Bs
        Hoja81.Range("BE15").FormulaR1C1 = "TOTAL" & Chr(10) & "Estimado" & Chr(10) & "Bs."
        Hoja81.Range("BE16").FormulaLocal = "=SI(AJ16=0;0;BB16*F16)"
        'Hoja81.Range("BF15").FormulaR1C1 = "TOTAL" & Chr(10) & "Estimado" & Chr(10) & "Bs."
        
        'SUMAS TOTALES
        Hoja81.Range("BB6") = "TOTAL EN Bs."
        Hoja81.Range("BE6").FormulaLocal = "=SUMA(BE15:BE17)"
        Hoja81.Range("BB7") = "TOTAL EN $"
        Hoja81.Range("BE7").FormulaLocal = "=SUMA(BC15:BC17)"
        Hoja81.Range("BE7").Select
        Selection.NumberFormat = _
        "_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* ""-""??_);_(@_)"
        'Hoja81.Range("BF7") = ""
    '   Total en Bs para la hoja de Control Mensual
        Hoja81.Range("BK1").FormulaLocal = "=BE6"
    
        'HOJA RESUMEN
        Hoja1.Range("N3").FormulaLocal = "=TEXTO('Estimado Base'!BA16;""#.##0,00"")"
        
        'FORMATO HOJA
        Worksheets("Estimado Base").Rows("10:13").Hidden = True
 
        
    'FORMATO DE CELDAS
    Hoja81.Range("BE13:BE14").Select
    Selection.ClearContents
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Interior.ColorIndex = xlNone
    
    Hoja81.Range("BB8:BE11").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Interior.ColorIndex = xlNone
    Selection.ClearContents
    
    
    Range("AG15").Select
    Selection.Copy
    Range("AB15").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("AA14:AF14").Select
    Selection.ClearContents
    Range("AA14:AL14").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.UnMerge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    
    Hoja81.Range("B16").Select
    
    Columns("AA:AA").EntireColumn.Hidden = True
    Columns("AC:AF").EntireColumn.Hidden = True
    
    End If
               
        

       
    'Empieza numeración
    FILAS = TextBox5 + 15
    Rows("16:16").Select
    Selection.Copy
    Range("A16") = 1
    
    'INGRESA CANTIDAD DE FILAS
    If TextBox5 = 1 Then '1 renglón
        GoTo asd
    End If
    If TextBox5 = 2 Then '2 renglones
        Rows("17:17").Select
        Selection.Insert Shift:=xlDown
        Application.CutCopyMode = False
        Range("A17") = 2
        GoTo asd
    End If
    
    Rows("17:" & FILAS).Select 'Más de 2 renglones
    Selection.Insert Shift:=xlDown
    Application.CutCopyMode = False
    Range("A17") = 2
    Range("A16:A17").Select
    Selection.AutoFill Destination:=Range("A16:A" & FILAS), Type:=xlFillDefault
asd:
    Application.CutCopyMode = False
       
    'DATOS GENERALES
    Range("E3") = Trim(TextBox10)
    Range("E4") = TextBox3 'Descripción
    Range("BE4") = TextBox5 'Renglones
    If CheckBox1 = True Then
        Range("E7") = TextBox8 'Validez Menor
    End If
    If CheckBox2 = True Then
        Range("E7") = TextBox12 'Validez Mayor
    End If
    Range("E8") = TextBox11 'Tiempo de Entrega
    Range("E6") = TextBox9  'Solicitante
   
    'DATOS GENERALES
    Range("BF1") = "1"
    If Range("BF3") <> "1" Then
    MsgBox "Debe Indicar datos numéricos en Tiempo de Validez/Entrega para continuar", vbCritical, "Advertencia"
    GoTo dfr
    End If
    UserForm4.Hide 'Oculta Formulario
dfr:

'AGREGAR FILAS HOJA RESUMEN

    Application.ScreenUpdating = False
    HojaActiva = ActiveSheet.Name
    Dim filx As Double
    Hoja1.Range("R3").FormulaLocal = "=AHORA()"
    Hoja1.Range("S3").FormulaLocal = "=CELDA(""FILENAME"")"
    Hoja1.Range("T3").FormulaLocal = "=SI(K3=0;CONCATENAR(M3;"" "";TEXTO(L3;""#.##0,00""));CONCATENAR(""Bs "";TEXTO(K3;""#.##0,00"")))"
  
    filx = TextBox5 + 2
    If filx > 2 Then
    Application.GoTo Worksheets("RESUMEN").Range("A3:U3")
    Selection.AutoFill Destination:=Worksheets("RESUMEN").Range("A3:U" & filx), Type:=xlFillDefault
    Application.CutCopyMode = False
    End If
    
xdt:
    Worksheets("Estimado Base").Activate

End Sub

Private Sub CommandButton2_Click()
'BuscarProceso()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim Fila As String
Dim BUSCAR As String
Dim CARPETA As String
Dim ARCHIVO As String
Dim DIRECCION As Excel.Worksheet

If CreateObject("WScript.Network").UserName = "REISEL SANCHEZ" Then
    CARPETA = "D:\Desktop\CONTROL MENSUAL\"
    GoTo GHG
End If

If CreateObject("WScript.Network").UserName = "VARGASLY" Then
    CARPETA = "Z:\EDC\BASE DATOS\"
Else
    CARPETA = "H:\EDC\BASE DATOS\"
End If
GHG:

    ARCHIVO = "CONTROL MENSUAL.xls"
   
    On Error Resume Next
    Workbooks(ARCHIVO).Activate
    If Err = 0 Then
        GoTo RRR
    End If
        
    Err.Clear 'Clear erroneous errors
    Workbooks.Open CARPETA & ARCHIVO, ReadOnly:=True

RRR:
ThisWorkbook.Activate
BUSCAR = Trim(TextBox10)   'Numero de Referencia
NFila = Workbooks(ARCHIVO).Worksheets("CONTROL").Range("A:A").Find(What:=BUSCAR, lookat:=xlWhole).Row
If IsEmpty(NFila) = True Then
    MsgBox ("NO SE ENCONTRO EL CORRELATIVO EN EL CONTROL MENSUAL")
    GoTo HJK
End If

Set DIRECCION = Workbooks(ARCHIVO).Worksheets("CONTROL")

'TITULO
TextBox3 = DIRECCION.Cells(NFila, 11)
'CANT POSICIONES
TextBox5 = DIRECCION.Cells(NFila, 13)
'DIAS OFERTAS MENOR
TextBox8 = DIRECCION.Cells(NFila, 17)
'DIAS OFERTAS MAYOR
TextBox12 = DIRECCION.Cells(NFila, 18)
'TIEMPO DE ENTREGA
TextBox11 = DIRECCION.Cells(NFila, 19)
'SOLICITANTE
TextBox9 = DIRECCION.Cells(NFila, 8)
'TIPO DE PROCESO
If DIRECCION.Cells(NFila, 5) = "ES" And DIRECCION.Cells(NFila, 14) = "VEF" Then
    OptionButton6 = True
    OptionButton8 = False
    OptionButton9 = False
End If
If DIRECCION.Cells(NFila, 5) = "ES" And DIRECCION.Cells(NFila, 14) = "MIXTO" Then
    OptionButton6 = False
    OptionButton9 = True
    OptionButton8 = False
End If
If DIRECCION.Cells(NFila, 5) = "EX" Then
    OptionButton6 = False
    OptionButton9 = False
    OptionButton8 = True
End If

HJK:
Workbooks(ARCHIVO).Close

End Sub


Private Sub OptionButton6_Click()
    OptionButton8 = False
    OptionButton9 = False
End Sub

Private Sub OptionButton8_Click()
    OptionButton6 = False
    OptionButton9 = False
End Sub

Private Sub OptionButton9_Click()
    OptionButton6 = False
    OptionButton8 = False
End Sub
Private Sub UserForm_Activate()
    On Error Resume Next
   
    If Range("BF1") = "1" Then
        GoTo a
    End If

    Hoja81.Range("BE1").Select
    ActiveCell.FormulaR1C1 = "=+TODAY()" 'Formula dia de hoy
    Selection.Copy      'Copia formula
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=False 'Pega valor de la formula
    Application.CutCopyMode = False
    Hoja81.Range("B16").Select
    
    TextBox6 = Range("BE1")

    Range("E4") = "" 'Descripción
    Range("E7") = "" 'Validez
    Range("E8") = "" 'Tiempo de Entrega
    Range("BE4") = "" 'Renglones
    Range("E6") = "" 'Solicitante
    Range("E3") = "" 'Referencia Estimado
a:

    TextBox10.SetFocus

asc:
End Sub

Private Sub UserForm_Deactivate()
    On Error Resume Next
    If Range("BF1") <> "1" Then
        MsgBox "Debe completar los datos para continuar", vbCritical, "Advertencia"
        UserForm4.Show
    End If
End Sub


