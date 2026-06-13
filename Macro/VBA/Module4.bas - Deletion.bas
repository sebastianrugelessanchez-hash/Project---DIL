Attribute VB_Name = "Module4"
Sub nuevo_libros()
Attribute nuevo_libros.VB_ProcData.VB_Invoke_Func = " \n14"
Application.DisplayAlerts = False
RUTA = ThisWorkbook.Path
FECHA = Format(Now, "mm.dd.yyyy")
    Sheets(Array("AUDIT", "PIVOT")).Select
    Sheets("PIVOT").Activate
    Sheets(Array("AUDIT", "PIVOT")).Copy
    ActiveWorkbook.SaveAs Filename:=RUTA & "\WF Report FINAL " & FECHA & ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
    Sheets("MACRO").Select
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    Windows("ERROR LOG 09.11.2022.XLSX").Activate
    Sheets("Sheet1").Select
    Sheets("Sheet1").Copy After:=Workbooks("VL06G 09.11.2022.XLSX").Sheets(1)
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "ERROR LOG"
    Windows("ERROR LOG 09.11.2022.XLSX").Activate
    ActiveWindow.Close
    Windows("AUDIT MACRO.xlsm").Activate
End Sub

Sub CerraLibros()
Dim cwb As Workbook
For Each cwb In Workbooks
    If cwb.Name <> ThisWorkbook.Name Then
    cwb.Close SaveChanges:=True
    End If
Next cwb
End Sub
