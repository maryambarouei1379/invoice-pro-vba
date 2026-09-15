Attribute VB_Name = "Module1"
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'

'
    Range("A1:I25").Select
    Selection.Copy
    Sheets("Sheet5").Select
    Range("A440").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(3, 0).Range("A1").Select
    ActiveSheet.Paste
    Sheets("Sheet3").Select
    Application.CutCopyMode = False
    Range("B11:I20").Select
    Selection.ClearContents
    Range("C7:F7,H7:I7,C8:I8,H1:I1,H2:I2").Select
    Range("H2").Activate
    Selection.ClearContents
    Range("H1:I1").Select
End Sub
