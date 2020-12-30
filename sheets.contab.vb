Private Sub CommandButton1_Click()
  Dim hoja As Worksheet
  
  Sheets("CONTABILIZADOS").Range("C4").ClearContents
Sheets("CONTABILIZADOS").Range("C4").ClearFormats

Sheets("CONTABILIZADOS").Range("A9:AT5000").ClearContents
Sheets("CONTABILIZADOS").Range("B4").ClearContents
Sheets("CONTABILIZADOS").Range("E1").ClearContents
Sheets("VALIDACION").Range("A3:E5000").ClearContents

For Each hoja In ThisWorkbook.Worksheets
    If hoja.FilterMode Then
        hoja.ShowAllData
    End If
Next hoja



Sheets("ACTIVOS").Range("A2:BZ5000").ClearContents
Sheets("CANCELADOS").Range("A2:AZ5000").ClearContents
Sheets("VALIDA CANCELADOS").Range("A2:AV5000").ClearContents
Sheets("PAGOS").Range("A2:K5000").ClearContents
Sheets("PAGOS").Range("N2:AV5000").ClearContents
Sheets("VARIACION").Range("A4:AV5000").ClearContents
Sheets("Restructurados").Range("A2:AV5000").ClearContents
Sheets("NOVEDADES").Range("A4:AB5000").ClearContents

End Sub
Private Sub CommandButton2_Click()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False
    Call abrirContabilizados 'pega contabilizados
    Call Macro2
    'Call ulcd
    
    
If Trim(Sheets("CONTABILIZADOS").Range("B4")) = "MSO" Then
   MsgBox ("Por favor si va a realizar Mso educadores al momento de realizar el reporte con la macro en la casilla (B4) poner MSOEDU"), vbCritical, "Informativo"
ElseIf Trim(Sheets("CONTABILIZADOS").Range("B4")) = "GCA" Then
MsgBox ("No olvidar pasar las cedulas de adm a salud"), vbCritical, "Informativo"
End If

   ' MsgBox "Contabilizados Listo", vbInformation
    
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
End Sub

Private Sub CommandButton3_Click()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False
CONVENIO = Sheets("CONTABILIZADOS").Cells(4, 2)

For Each hoja In ThisWorkbook.Worksheets
    If hoja.FilterMode Then
        hoja.ShowAllData
    End If
Next hoja
'se debe enlazar a la sheets de la macro
Call GuardarValida

Application.Speech.Speak "Valida Guardado.. op.plus..mas y mejor", True
MsgBox ("Valida- " & CONVENIO & " -Guardado"), vbInformation, "Macro Opplus"
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
End Sub

