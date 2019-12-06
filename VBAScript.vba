Sub homework()
'Hacerlo para todas las worksheets
'Dim ws As Worksheet
For Each ws In Worksheets

'Variables primera parte
    Dim i As Long
    Dim lastrow As Long
    Dim yopen As Double
    Dim yclose As Double
    'Dim contador As Integer
    Dim ychange As Double
    'Dim ychangepercentage As Double
    Dim tickName As String
    'Dim volumen As Long
    'Dim totalVol As Long
    'Dim tickeName() As String
'Variables para obtener maximos y minimos
    Dim Summary_Table_Row As Integer
    Dim tickMaxIncrease As String
    Dim tickMinDecrease As String
    Dim tickMaxVolume As String
    Dim Maxincrease As Double
    Maxincrease = 0
    Dim minDecrease As Double
    minDecrease = 0
    Dim maxVolume As Double
    maxVolume = 0
'Encabezados de variables
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "% Yearly Change"
    ws.Range("L1") = "Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest % increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest total volume"
    
    

    Summary_Table_Row = 2
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'contador = 0
    yopen = ws.Cells(2, 3).Value
    For i = 2 To lastrow
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
               tickeName = ws.Cells(i, 1).Value
                'Calculando el valor de year change
                    yclose = ws.Cells(i, 6).Value
                    ychange = yclose - yopen
                    
                    'If ychangepercentage division entre 0
                    If yopen <> 0 Then
                    ychangepercentage = (ychange / yopen) * 100
                    Else
                    MsgBox ("Error al dividir por 0")
                    End If
                    'If Cambio de formato
                    If ychangepercentage <= 0 Then
                        ws.Range("k" & Summary_Table_Row).Interior.ColorIndex = 3
                    ElseIf ychangepercentage > 0 Then
                    ws.Range("k" & Summary_Table_Row).Interior.ColorIndex = 4
                    End If
                    totalVol = totalVol + ws.Cells(i, 7).Value
                ws.Range("I" & Summary_Table_Row).Value = tickeName
                ws.Range("j" & Summary_Table_Row).Value = ychange
                ws.Range("k" & Summary_Table_Row).Value = (CStr(ychangepercentage) & "%")
               ws.Range("l" & Summary_Table_Row).Value = totalVol
                Summary_Table_Row = Summary_Table_Row + 1
                yopen = ws.Cells(i + 1, 3).Value
                ychange = 0
                'Encontrar el valor maximo minino
                If ychangepercentage > Maxincrease Then
                    Maxincrease = ychangepercentage
                    tickMaxIncrease = tickeName
                ElseIf ychangepercentage < minDecrease Then
                    minDecrease = ychangepercentage
                    tickMinDecrease = tickeName
                End If
                
                If totalVol > maxVolume Then
                    maxVolume = totalVol
                    tickMaxVolume = tickeName
                End If
                
                'Reset variables por ticker
                    ychangepercentage = 0
                    totalVol = 0
                
            Else
            totalVol = totalVol + ws.Cells(i, 7).Value
            
            End If
           
    Next i

                ws.Range("Q2").Value = (CStr(Maxincrease) & "%")
                ws.Range("Q3").Value = (CStr(minDecrease) & "%")
                ws.Range("P2").Value = tickMaxIncrease
                ws.Range("P3").Value = tickMinDecrease
                ws.Range("Q4").Value = maxVolume
                ws.Range("P4").Value = tickMaxVolume
Next ws


End Sub
