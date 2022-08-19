Private Sub CommandButton1_Click()
Unload About
End Sub

Private Sub Label4_Click()
On Error GoTo NoSuchFile
ActiveWorkbook.FollowHyperlink Address:="G:\Dcs\QMS_EHS_processes\40_Supply_Chain_processes\Incoming processes\3_ROM-Nr.7121 Mod de utilizare fisier Report Register rev.001.doc", NewWindow:=True
Exit Sub
NoSuchFile:
MsgBox ("Fisierul '3_ROM-Nr.7121 Mod de utilizare fisier Report Register rev.001' nu exista sau a fost mutat.")
End Sub
