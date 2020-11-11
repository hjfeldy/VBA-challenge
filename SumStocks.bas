Attribute VB_Name = "Module1"
'LUDICROUS MODE IS NOT MY CODE; I GOT IT FROM REDDIT
'username ViperSRT3g
'LudicrousMode does not affect my code in any way besides making it run faster
Public Sub LudicrousMode(ByVal Toggle As Boolean)
    Application.ScreenUpdating = Not Toggle
    Application.EnableEvents = Not Toggle
    Application.DisplayAlerts = Not Toggle
    Application.EnableAnimations = Not Toggle
    Application.DisplayStatusBar = Not Toggle
    Application.PrintCommunication = Not Toggle
    Application.Calculation = IIf(Toggle, xlCalculationManual, xlCalculationAutomatic)
End Sub

Sub stocks()

Call LudicrousMode(True)

For Each ws In Worksheets

    'boring setup
    Sheets(ws.Name).Select
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest Vol"
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Change"
    Range("K1").Value = "Pct Change"
    Range("L1").Value = "Total Volume"
    Dim theBeg As Double
    Dim theEnd As Double
    Dim volTot As Double
    Dim outCount As Integer
    outCount = 1
    Dim bigVol As Double
    Dim bigChange As Double
    Dim smallChange As Double
    bigVol = 0
    bigChange = 0
    smallChange = 0
    
    'Find number of rows to iterate over
    RowCount = 1
    While Cells(RowCount, 1).Value <> ""
        RowCount = RowCount + 1
    Wend
    RowCount = RowCount - 1
    
    'Declare totalVolume and beginningPrice vars
    volTot = Cells(2, 7).Value
    theBeg = Cells(2, 3).Value
    
    For i = 2 To RowCount
    
        volTot = volTot + Cells(i, 7).Value
        'If this cell doesn't equal the next then
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
            'increment outputRow
            outCount = outCount + 1
            
            'fill output-table with data
            Cells(outCount, 9).Value = Cells(i, 1).Value
            Cells(outCount, 12).Value = volTot
            theEnd = Cells(i, 6).Value
            Change = theEnd - theBeg
            Cells(outCount, 10).Value = Change
            
            'Conditional Formatting
            If Change > 0 Then
                Cells(outCount, 10).Interior.ColorIndex = 4
            Else
                Cells(outCount, 10).Interior.ColorIndex = 3
            End If
            
            'Fix empty data -- otherwise we're dividing by zero
            If theBeg = 0 Then
                theBeg = theEnd + 0.01 'ensure theBeg is never 0, and that if it is, we record a 0 in the output-table for our faulty data
                Cells(outCount, 11).Value = 0
            Else
                Cells(outCount, 11).Value = Change / theBeg
            End If
            
            'Check for biggest/smallest change and biggest trading volume
            If (Change / theBeg) > bigChange Then
                bigChange = (Change / theBeg)
                bigChangeStock = Cells(i, 1).Value
                Yr = Left(Cells(i, 2).Value, 4)
            End If
            If (Change / theBeg) < smallChange Then
                smallChange = (Change / theBeg)
                smallChangeStock = Cells(i, 1).Value
                Yr = Left(Cells(i, 2).Value, 4)
            End If
            If volTot > bigVol Then
                bigVol = volTot
                bigVolStock = Cells(i, 1).Value
                Yr = Left(Cells(i, 2).Value, 4)
            End If
            
            'format %
            Cells(outCount, 11).NumberFormat = "0.00%"
            
            'Reset vars before going on to next stock
            theBeg = Cells(i + 1, 3).Value
            volTot = 0
        
        End If
    Next i
    
    
Next ws

'Fill in bonus output table (fill every individual sheet, but the data is global)
For Each ws In Worksheets
    Sheets(ws.Name).Select
    Range("N2").Value = "Biggest % Change"
    Range("N3").Value = "Smallest % Change"
    Range("N4").Value = "Biggest Trading Volume"
    
    Range("O2").Value = bigChangeStock
    Range("P2").Value = bigChange
    Range("P2").NumberFormat = "0.00%"
    Range("O4").Value = bigVolStock
    Range("P4").Value = bigVol
    Range("O3").Value = smallChangeStock
    Range("P3").Value = smallChange
Next ws

Call LudicrousMode(False)
End Sub


