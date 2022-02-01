Attribute VB_Name = "Módulo11"
Sub Facturas_B()
Sheets("Sheet1").Select      'nombre de la hoja con la información
col = "B"                   'columna para aplicar la condición
'texto de la condición
'Para una fecha: "10/07/2017" el formato debe ser dd/mm/aaaa
'Para un número: "123"
texto = "6 - Factura B"    '
valor = texto
If IsNumeric(texto) Then valor = Val(texto)
If IsDate(texto) Then valor = CDate(texto)    '
Application.ScreenUpdating = False
For i = Range(col & Rows.Count).End(xlUp).Row To 1 Step -1
If LCase(Cells(i, "B")) = LCase(valor) Then
Rows(i).Delete
End If
Next
Application.ScreenUpdating = True
MsgBox "Facturas B eliminadas", vbInformation, "Juani"
End Sub
Sub Comprobantes_Jaque()
Sheets("Sheet1").Select      'nombre de la hoja con la información
col = "E"                   'columna para aplicar la condición
'texto de la condición
'Para una fecha: "10/07/2017" el formato debe ser dd/mm/aaaa
'Para un número: "123"
texto = "SI"    '
valor = texto
If IsNumeric(texto) Then valor = Val(texto)
If IsDate(texto) Then valor = CDate(texto)    '
Application.ScreenUpdating = False
For i = Range(col & Rows.Count).End(xlUp).Row To 1 Step -1
If LCase(Cells(i, "E")) = LCase(valor) Then
Rows(i).Delete
End If
Next
Application.ScreenUpdating = True
MsgBox "Comprobantes en Jaque eliminados", vbInformation, "Juani"
End Sub
Sub Notas_Credito()
Sheets("Sheet1").Select      'nombre de la hoja con la información
col = "B"                   'columna para aplicar la condición
'texto de la condición
'Para una fecha: "10/07/2017" el formato debe ser dd/mm/aaaa
'Para un número: "123"
texto = "3 - Nota de Crédito A"    '
valor = texto
If IsNumeric(texto) Then valor = Val(texto)
If IsDate(texto) Then valor = CDate(texto)    '
Application.ScreenUpdating = False
For i = Range(col & Rows.Count).End(xlUp).Row To 1 Step -1
If LCase(Cells(i, "B")) <> LCase(valor) Then
Rows(i).Delete
End If
Next
Application.ScreenUpdating = True
MsgBox "Notas de Crédito extraídas", vbInformation, "Juani"
End Sub
