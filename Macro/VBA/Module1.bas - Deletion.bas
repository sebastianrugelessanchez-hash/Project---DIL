Attribute VB_Name = "Module1"
Sub DESCARGA_US()


Application.DisplayAlerts = False

   Set SapGuiAuto = GetObject("SAPGUI")
   Set App = SapGuiAuto.GetScriptingEngine
   Set Connection = App.Children(0)
   Set session = Connection.Children(0)
   
   
fecha1 = Range("d6").Value
fecha2 = Range("E6").Value
RUTA = ThisWorkbook.Path
FECHA = Format(Now, "mm.dd.yyyy")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/NZWF_AUDIT"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtV-LOW").Text = "BILLING US ACM"
session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
session.findById("wnd[1]").sendVKey 0
'session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtP_DATE-LOW").Text = fecha1
session.findById("wnd[0]/usr/ctxtP_DATE-HIGH").Text = fecha2
session.findById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").press

Sheets("MACRO").Select
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[24]").press

session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell 0, "SOLDTO"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/cmbG_LISTBOX").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = RUTA & "\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "AUDIT " & FECHA & ".xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]").sendVKey 12
session.findById("wnd[0]").sendVKey 12

libro = "AUDIT " & FECHA & ".xlsx"

    Workbooks.Open Filename:= _
        RUTA & "\" & libro


    Windows("AUDIT MACRO.xlsm").Activate
    Sheets("AUDIT").Select
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    
    Windows(libro).Activate
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Columns("O:O").Select
    Selection.Insert Shift:=xlToRight
    Range("A1000000").Select
    Selection.End(xlUp).Select
i = ActiveCell.Row
    Range("A2:Y" & i).Select
    Selection.Copy
    Windows("AUDIT MACRO.xlsm").Activate
    Range("A1000000").Select
    Selection.End(xlUp).Select
p = ActiveCell.Row + 1
    Range("A" & p).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Range("B342").Activate
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("D:D").Select
    Range("D342").Activate
    Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("F:F").Select
    Range("F342").Activate
    Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("G:G").Select
    Range("G342").Activate
    Selection.TextToColumns Destination:=Range("G1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("I:I").Select
    Range("I342").Activate
    Selection.TextToColumns Destination:=Range("I1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("K:K").Select
    Range("K342").Activate
    Selection.TextToColumns Destination:=Range("K1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("L:L").Select
    Range("L342").Activate
    Selection.TextToColumns Destination:=Range("L1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("M:M").Select
    Range("M342").Activate
    Selection.TextToColumns Destination:=Range("M1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

    Columns("P:P").Select
    Range("P327").Activate
    Selection.TextToColumns Destination:=Range("P1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("Q:Q").Select
    Range("Q327").Activate
    Selection.TextToColumns Destination:=Range("Q1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("S:S").Select
    Range("S327").Activate
    Selection.TextToColumns Destination:=Range("S1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Windows(libro).Activate
    ActiveWindow.Close
    
    Sheets("AUDIT").Select
    Range("A100000").Select
    Selection.End(xlUp).Select
q = ActiveCell.Row
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=+VLOOKUP(RC[-1],'SALES OFFICE'!R1C3:R264C4,2,0)"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C" & q)
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=+VLOOKUP(RC[-1],'SALES GROUP'!R1C2:R325C4,3,0)"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E" & q)
    Columns("C:C").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("E:E").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("C1").Select
    Range("H2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(LEFT(RC[-1],1)=""4"",""RMX"",IF(LEFT(RC[-1],1)=""3"",""AGG"",IF(LEFT(RC[-1],1)=""8"",""ASP"")))"
    Range("H2").Select
    Selection.AutoFill Destination:=Range("H2:H" & q)
    Columns("H:H").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("H1").Select

Call validaciones_error
Call inventary
Call nuevo_libros
Call CerraLibros

End Sub

Sub validaciones_error()
Application.DisplayAlerts = False

    Sheets("AUDIT").Select
    Columns("N:N").Select
    Selection.AutoFilter
    Sheets("validation").Select
    Range("B1").Select
    Selection.End(xlDown).Select
i = ActiveCell.Row - 1
p = 2
For w = 1 To i
Sheets("validation").Select
Errores = Range("b" & p).Value
Action = Range("c" & p).Value
    Sheets("AUDIT").Select
    ActiveSheet.Range("$N$1:$N$1000000").AutoFilter Field:=1, Criteria1:= _
        "=*" & Errores & "*", Operator:=xlAnd
    Range("N1000000").Select
    Selection.End(xlUp).Select
q = ActiveCell.Row

If q = 1 Then
GoTo siguee
End If
    Range("O2:O" & q).Select

    Selection.FormulaR1C1 = Action
siguee:
p = p + 1

Next
Selection.AutoFilter

    
End Sub



Sub inventary()
Attribute inventary.VB_ProcData.VB_Invoke_Func = " \n14"
Application.DisplayAlerts = False

RUTA = ThisWorkbook.Path
FECHA = Format(Now, "mm.dd.yyyy")

    Columns("A:A").Select
    Selection.AutoFilter
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$A$1000000").AutoFilter Field:=1, Criteria1:= _
        "=*Ticket not Goods Issued*", Operator:=xlAnd
    Range("A1:W1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("INVENTORY").Select
    Range("A1").Select
    ActiveSheet.Paste
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A9").Select
    Sheets(Array("INVENTORY", "INVENTORY PIVOT")).Select
    Sheets("INVENTORY PIVOT").Activate
    Sheets(Array("INVENTORY", "INVENTORY PIVOT")).Copy
    ActiveWorkbook.SaveAs Filename:= _
        RUTA & "\Inventory " & FECHA & ".xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False

    Windows("AUDIT MACRO.xlsm").Activate
    
    Sheets("pivot").Select
        Sheets("INVENTORY").Select

    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Sheets("AUDIT").Select
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.ShowAllData
    Range("N1").Select
    Sheets("PIVOT").Select
    Range("B4").Select
    ActiveWorkbook.RefreshAll
End Sub
Sub Macro21()
Attribute Macro21.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro21 Macro
'

'

End Sub
