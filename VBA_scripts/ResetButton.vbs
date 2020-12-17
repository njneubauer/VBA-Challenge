Attribute VB_Name = "ResetButton"
Sub RestButton_Click()
For Each ws In Worksheets
    ws.Activate
    Range("J1", Range("M1").End(xlDown)).ClearContents
    Range("O1", Range("Q1").End(xlDown)).ClearContents
    Range("J1", Range("M1").End(xlDown)).ClearFormats
Next ws

Sheet1.Activate

End Sub

