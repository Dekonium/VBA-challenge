Attribute VB_Name = "Module1"
Sub Worsksheets()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call WallStreet
    Next
    Application.ScreenUpdating = True
End Sub



Sub WallStreet()

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Dim Ticker As String
Dim Change As Double
Dim ChangePerc As Double
Dim Volume As Double
Dim Rown As Integer
Dim CountRows As Double
Rown = 2
CountRows = 0
Volume = 0
For i = 2 To 800000

    If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
        CountRows = CountRows + 1
        Volume = Volume + Cells(i, 7)
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Range("I" & Rown).Value = Ticker
        
        Change = Change + (Cells(i, 6) - Cells(i - CountRows, 3))
        ChangePerc = Change / Cells(i - CountRows, 3)
        
        Range("J" & Rown).Value = Change
        Range("K" & Rown).Value = ChangePerc
        Range("K" & Rown).NumberFormat = "0.00%"
        
        Range("L" & Rown).Value = Volume
        
        Rown = Rown + 1
        CountRows = 0
        Volume = 0
        Change = 0
    End If
    
Next i

For j = 2 To Rown - 1

    If Cells(j, 10).Value < 0 And Not IsEmpty(Cells(j, 10).Value) Then
        Cells(j, 10).Interior.ColorIndex = 3
    Else
        Cells(j, 10).Interior.ColorIndex = 4
    
    End If
Next j

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).NumberFormat = "0.00%"

Dim Increase As Double
Dim Decrease As Double
Dim TotalV As Double

Increase = 0
Decrease = 0
TotalV = 0

For h = 2 To Rown - 1
    If Cells(h, 11).Value > Increase Then
        Cells(2, 17).Value = Cells(h, 11)
        Increase = Cells(h, 11)
        
        Cells(2, 16).Value = Cells(h, 9)
        
    End If
Next h

For p = 2 To Rown - 1
    If Cells(p, 11).Value < Decrease Then
        Cells(3, 17).Value = Cells(p, 11)
        Decrease = Cells(p, 11)
        
        Cells(3, 16).Value = Cells(p, 9)
        
    End If
Next p

For y = 2 To Rown - 1
    If Cells(y, 12).Value > Volume Then
        Cells(4, 17).Value = Cells(y, 12)
        Volume = Cells(y, 12)
        
        Cells(4, 16).Value = Cells(y, 9)
        
    End If
Next y



End Sub

