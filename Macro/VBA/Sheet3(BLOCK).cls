VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()

Set SapGuiAuto = GetObject("SAPGUI")
Set App = SapGuiAuto.GetScriptingEngine
Set Connection = App.Children(0)
Set session = Connection.Children(0)

Dim OU As Workbook
Dim BLOCK As Worksheet
Set BLOCK = ThisWorkbook.Sheets("BLOCK")

For i = 2 To ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row

On Error GoTo error
'transaction
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nvl02n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = BLOCK.Range("A" & i).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\06").Select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\06/ssubSUBSCREEN_BODY:SAPMV50A:2112/cmbLIKP-FAKSK").Key = "03"
session.findById("wnd[0]/tbar[0]/btn[11]").press
Range("B" & i) = "Blocked"

error:
If (Err.Number > 0) Then
Range("B" & i) = "Not blocked"

On Error GoTo 0
On Error GoTo -1

End If

Next

End Sub

Private Sub CommandButton2_Click()

Set SapGuiAuto = GetObject("SAPGUI")
Set App = SapGuiAuto.GetScriptingEngine
Set Connection = App.Children(0)
Set session = Connection.Children(0)

Dim OU As Workbook
Dim BLOCK As Worksheet
Set BLOCK = ThisWorkbook.Sheets("BLOCK")

For i = 2 To ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row

On Error GoTo error
'transaction
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nvl02n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = BLOCK.Range("A" & i).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\06").Select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\06/ssubSUBSCREEN_BODY:SAPMV50A:2112/cmbLIKP-FAKSK").Key = " "
session.findById("wnd[0]/tbar[0]/btn[11]").press
Range("B" & i) = "Unblocked"

error:
If (Err.Number > 0) Then
Range("B" & i) = "Not Unblocked"

On Error GoTo 0
On Error GoTo -1

End If

Next
End Sub
