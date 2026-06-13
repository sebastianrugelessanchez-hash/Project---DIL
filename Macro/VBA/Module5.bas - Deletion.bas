Attribute VB_Name = "Module5"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Sheets(Array("INVENTORY", "INVENTORY PIVOT")).Select
    Sheets("INVENTORY PIVOT").Activate
    Sheets(Array("INVENTORY", "INVENTORY PIVOT")).Copy
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Sheets("INVENTORY").Select
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Range("B4:C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("Copy of VL06G 11.01.2023.XLSX").Activate
    Sheets.Add After:=ActiveSheet
    Range("D5").Select
    ActiveSheet.Paste
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("G16").Select
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "PIVOT"
End Sub
