VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StkExch()
    Dim Ws As Worksheet
    'loop thru every worksheet and execute the following code
    For Each Ws In Worksheets
    Dim tick As String
    Dim opn As Double
    Dim clz As Double
    Dim chng As Double
    Dim vol As Variant
    
    Dim lr As Double
    Dim t As Double
    Dim newtick As String
    Dim pctchng As Double

        Application.ScreenUpdating = False
         lr = Ws.Cells(Rows.Count, 1).End(xlUp).Row

            'setting the labels for the table to be created
            Ws.Range("K1:S1").Font.Bold = True
            
            Ws.Range("Q2:Q4").Font.Bold = True
            Ws.Range("K1").Value = "TICKER"
            Ws.Range("L1").Value = "YEARLY CHANGE"
            Ws.Range("M1").Value = "PERCENT CHANGE"
            Ws.Range("N1").Value = "TOTAL STOCK VOLUME"
            Ws.Range("R1").Value = "TICKER"
            Ws.Range("S1").Value = "VALUE"
            Ws.Range("Q2").Value = "GREATEST % INCREASE"
            Ws.Range("Q3").Value = "GREATEST % DECREASE"
            Ws.Range("Q4").Value = "GREATEST TOTAL VOLUME"

            'not required but inserting a difinitive break between the parent data and the result tables
            Ws.Range("I1:I" & lr).Interior.Color = RGB(234, 105, 34)
            Ws.Columns.AutoFit
            



            
    'Clear out data from last run
             Ws.Range(Ws.Cells(2, 11), Ws.Cells(lr + 1, 14)).ClearContents
             Ws.Range(Ws.Cells(2, 11), Ws.Cells(lr + 1, 14)).ClearFormats
             Ws.Range(Ws.Cells(2, 18), Ws.Cells(4, 19)).ClearContents
             
             
             'initializing variable values ahead of loop
             vol = Ws.Cells(2, 7)
             newtick = Ws.Cells(2, 1)
              opn = Ws.Cells(2, 3).Value
              
                 'new table row variable
                 t = 2
    'Lets Loop!
              For i = 2 To lr
                               
                  clz = Ws.Cells(i, 6).Value
                  
                  
                  
                  ' Checking one row ahead for new ticker value,
                    
                   If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
                   
                   'if condition is true we assign values to the new table
                   chng = opn - clz
                   
                    ' !!! here is the bug fix avoid dividing by 0!!!
                   If opn <> 0 Then
                   pctchng = Application.WorksheetFunction.Round(chng / opn, 5) * 100
                   Else: pctchng = 0
                   End If
                   Ws.Cells(t, 13) = pctchng
                    Ws.Cells(t, 11) = newtick
                   
                    Ws.Cells(t, 12) = chng
                    
                    'nested if to satisfy the color requirement for the change column
                    If chng < 0 Then
                   Ws.Cells(t, 11).Interior.Color = vbRed
                   Else: Ws.Cells(t, 11).Interior.Color = vbGreen
                   End If
                   Ws.Cells(t, 14).Value = vol
                   
                   'after the table values are assigned, we increment the table row and other variables are reset
                     t = t + 1
                      newtick = Ws.Cells(i + 1, 1)
                    vol = Ws.Cells(i + 1, 7).Value
                    opn = Ws.Cells(i + 1, 3)
                    
                    'original if condition not met continue to count up volume
                    Else
                     vol = Ws.Cells(i + 1, 7) + vol
                    End If
                    
                    
                    
              Next i
              
  'set variables for the max min  table
    Dim pctinc As Double
    Dim tickinc As String
    Dim pctdec As Double
    Dim tickdec As String
    Dim totalvol As Variant
    
    Dim tickvol As String

    'reset last row to the last row of the results table and assign arbitrary values ahead of the loop
    lr = Ws.Cells(Rows.Count, 11).End(xlUp).Row
    pctdec = -0.00001
    pctinc = 0.0001
    totalvol = 0
    
         'loop thru the table and check for greatest and least values
        For i = 2 To lr
        
           If Ws.Cells(i, 13).Value < 0 And Ws.Cells(i, 13).Value < pctdec Then
           pctdec = Ws.Cells(i, 13).Value
           tickdec = Ws.Cells(i, 11).Value
           ElseIf Ws.Cells(i, 13).Value >= 0 And Ws.Cells(i, 13).Value > pctinc Then
           pctinc = Ws.Cells(i, 13).Value
           tickinc = Ws.Cells(i, 11).Value
           End If
           If Ws.Cells(i, 14).Value > totalvol Then
           totalvol = Ws.Cells(i, 14).Value
           tickvol = Ws.Cells(i, 11).Value

    End If
    Next i
             'assign the values to the max min table
             Ws.Cells(2, 18).Value = tickinc
             Ws.Cells(2, 19).Value = pctinc
             Ws.Cells(3, 18).Value = tickdec
             Ws.Cells(3, 19).Value = pctdec
             Ws.Cells(4, 18).Value = tickvol
             Ws.Cells(4, 19).Value = totalvol
            
'off to the next worksheet
  Next Ws
  
            
                    
End Sub
