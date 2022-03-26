# SAP-GUI-Scripting
Automate Scripts to generate report in SAP




Sub SCP()

Dim Appl As Object
Dim Connection As Object
Dim session As Object
Dim WshShell As Object
Dim SapGui As Object
Dim pic As Object
'Of course change for your file directory
Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", 4
Set WshShell = CreateObject("WScript.Shell")

Do Until WshShell.AppActivate("SAP Logon ")
    Application.Wait Now + TimeValue("0:00:01")
Loop

Set WshShell = Nothing

Set SapGui = GetObject("SAPGUI")
Set Appl = SapGui.GetScriptingEngine
Set Connection = Appl.OpenConnection("  ", _
    True)
Set session = Connection.Children(0)

'if You need to pass username and password
session.FindById("wnd[0]/usr/txtRSYST-MANDT").Text = "    "
session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = "    "
session.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = "    "
session.FindById("wnd[0]/usr/txtRSYST-LANGU").Text = "EN"

If session.Children.Count > 1 Then

    answer = MsgBox("You've got opened SAP already," & _
"please leave and try again", vbOKOnly, "Opened SAP")

    session.FindById("wnd[1]/usr/radMULTI_LOGON_OPT3").Select
    session.FindById("wnd[1]/usr/radMULTI_LOGON_OPT3").SetFocus
    session.FindById("wnd[1]/tbar[0]/btn[0]").press

    Exit Sub

End If

session.FindById("wnd[0]").maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "st22"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]").HardCopy "C:\xampp\htdocs\st22.png", 1
session.FindById("wnd[0]/usr/btnTODAY").press
session.FindById("wnd[0]").HardCopy "C:\xampp\htdocs\st22d.png", 1
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/ndb01"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/shellcont[1]/shell/shellcont[1]/shell").hierarchyHeaderWidth = 117
session.FindById("wnd[0]").HardCopy "C:\xampp\htdocs\db01.png", 1
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nal08"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]").HardCopy "C:\xampp\htdocs\al08.png", 1
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nsm12"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/txtSEQG3-GUNAME").Text = "*"
session.FindById("wnd[0]/usr/txtSEQG3-GUNAME").SetFocus
session.FindById("wnd[0]/usr/txtSEQG3-GUNAME").caretPosition = 1
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]").HardCopy "C:\xampp\htdocs\sm12.png", 1
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nsm13"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[8]").press
session.FindById("wnd[0]").HardCopy "C:\xampp\htdocs\sm13.png", 1
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nsost"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]").HardCopy "C:\xampp\htdocs\sost.png", 1
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nsm51"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]").HardCopy "C:\xampp\htdocs\sm51.png", 1
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nsm66"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[13]").press
session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, "WP_TYPE_DISP"
session.FindById("wnd[0]").HardCopy "C:\xampp\htdocs\sm66.png", 1
session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "WP_TYPE_DISP"
session.FindById("wnd[0]/tbar[1]/btn[38]").press
session.FindById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "SPO"
session.FindById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 3
session.FindById("wnd[1]/tbar[0]/btn[0]").press
session.FindById("wnd[0]").HardCopy "C:\xampp\htdocs\sm66spo.png", 1
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nsp01"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/tabsTABSTRIP_BL1/tabpSCR1/ssub%_SUBSCREEN_BL1:RSPOSP01NR:0100/txtS_RQOWNE-LOW").Text = ""
session.FindById("wnd[0]/usr/tabsTABSTRIP_BL1/tabpSCR1/ssub%_SUBSCREEN_BL1:RSPOSP01NR:0100/ctxtS_RQCRED-LOW").Text = ""
session.FindById("wnd[0]/usr/tabsTABSTRIP_BL1/tabpSCR1/ssub%_SUBSCREEN_BL1:RSPOSP01NR:0100/ctxtS_RQCRED-LOW").SetFocus
session.FindById("wnd[0]/usr/tabsTABSTRIP_BL1/tabpSCR1/ssub%_SUBSCREEN_BL1:RSPOSP01NR:0100/ctxtS_RQCRED-LOW").caretPosition = 0
session.FindById("wnd[0]/tbar[1]/btn[8]").press
session.FindById("wnd[0]/usr").verticalScrollbar.Position = 16028
session.FindById("wnd[0]").HardCopy "C:\xampp\htdocs\sp01.png", 1
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nsm37"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/chkBTCH2170-SCHEDUL").Selected = False
session.FindById("wnd[0]/usr/chkBTCH2170-READY").Selected = False
session.FindById("wnd[0]/usr/chkBTCH2170-FINISHED").Selected = False
session.FindById("wnd[0]/usr/chkBTCH2170-ABORTED").Selected = True
session.FindById("wnd[0]/usr/chkBTCH2170-RUNNING").Selected = False
session.FindById("wnd[0]/usr/txtBTCH2170-USERNAME").Text = "*"
session.FindById("wnd[0]/usr/chkBTCH2170-ABORTED").SetFocus
session.FindById("wnd[0]/tbar[1]/btn[8]").press
session.FindById("wnd[0]").sendVKey 5
session.FindById("wnd[0]").sendVKey 48
session.FindById("wnd[0]").HardCopy "C:\xampp\htdocs\sm37cnl.png", 1
session.FindById("wnd[0]/tbar[0]/btn[3]").press
session.FindById("wnd[0]/usr/chkBTCH2170-ABORTED").Selected = False
session.FindById("wnd[0]/usr/chkBTCH2170-RUNNING").Selected = True
session.FindById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").Text = ""
session.FindById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").SetFocus
session.FindById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").caretPosition = 0
session.FindById("wnd[0]/tbar[1]/btn[8]").press
session.FindById("wnd[0]").HardCopy "C:\xampp\htdocs\sm37act.png", 1

End Sub
