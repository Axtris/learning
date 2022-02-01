Attribute VB_Name = "Módulo2"
Sub Mis_Comprobantes_2()
'
' Mis_Comprobantes_2 Macro
'
' Acceso Directo: ctrl+shift+m
'
    ActiveCell.FormulaR1C1 = _
        "=+IF(ISERROR(AND(VLOOKUP(RC[-1],Hoja1!R1C1:R10000C2,1,FALSE),VLOOKUP(RC[3],Hoja1!R1C3:R10000C4,1,FALSE))),""NO"",""SI"")"
    Range("E2").Select
    
End Sub
