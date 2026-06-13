Attribute VB_Name = "Módulo1"
Public Sub main()

Set SapGuiAuto = GetObject("SAPGUI")
Set App = SapGuiAuto.GetScriptingEngine
Set Connection = App.Children(0)
Set session = Connection.Children(0)

Dim Order_Cancellation As Workbook
Dim OrderDeletion As Worksheet
Set OrderDeletion = ThisWorkbook.Sheets("OrderDeletion")

For i = 2 To ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row

On Error GoTo error
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = OrderDeletion.Range("A" & i).Value
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 9
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/mbar/menu[0]/menu[11]").Select
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
Range("B" & i) = "Deleted"

error:
If (Err.Number > 0) Then
Range("B" & i) = "Not Deleted"

On Error GoTo 0
On Error GoTo -1

End If

Next

End Sub
