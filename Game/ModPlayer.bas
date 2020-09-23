Attribute VB_Name = "ModPlayer"
Sub WarpPlayer(Map As Integer, X As Integer, Y As Integer)
X = X - 5
Y = Y - 5
OffsetX = X
OffsetY = Y
OPENMAP Map
LoadNPCs
Refresh_Screen
End Sub



Sub RenderStat()
On Error Resume Next
FrmMain.shpHelth.Width = (2415 / 100) * PlayerHelth
FrmMain.shpStam.Width = (2415 / 100) * PlayerStam
End Sub
