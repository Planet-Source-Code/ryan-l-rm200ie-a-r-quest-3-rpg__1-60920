Attribute VB_Name = "ModScreen"
Public Sub Refresh_Screen()
'Define to X and Y vars X1,2 and Y1,2
'X,y(1) are used for looping through the map array and getting the tile X
'X,y(2) is used to show where to draw the screen
Dim X(1 To 2) As Integer
Dim Y(1 To 2) As Integer

'Loop from Your Possition (top Left) to top left + 10 more tiles making the visable map
'10 x 10 tiles
'320 x 320 pixels
For X(1) = OffsetX To OffsetX + 10
For Y(1) = OffsetY To OffsetY + 10

'This here blits the tiles from Tileset to MapScreen
BitBlt FrmMain.PicMap.hDC, X(2) * 32, Y(2) * 32, 32, 32, FrmMain.PicTiles.hDC, Map(X(1), Y(1)).X * 32, Map(X(1), Y(1)).Y * 32, vbSrcCopy

'Check if a block is at the spot if it is put a red O there
If Map(X(1), Y(1)).Blocked = 1 Then
BitBlt FrmMain.PicMap.hDC, X(2) * 32, Y(2) * 32, 32, 32, FrmMain.PicColl.hDC, 0, 0, vbSrcAnd
End If

If Map(X(1), Y(1)).WarpMap > 0 Then
BitBlt FrmMain.PicMap.hDC, X(2) * 32, Y(2) * 32, 32, 32, FrmMain.picWarp.hDC, 0, 0, vbSrcPaint
End If

'These are just for scrolling its like a Loop without the For
'This line Scrolls to the bottem of the screen
Y(2) = Y(2) + 1
Next Y(1)
'Then it gets set to 0 (back to the top)
Y(2) = 0
'Then It moves over 1 tile and starts again
X(2) = X(2) + 1
Next X(1)

'These just set the captions n the main window
FrmMain.oX.Caption = OffsetX
FrmMain.oY.Caption = OffsetY

FPS = FPS + 1
'Refresh the main screen without this you wont see anything on the display
FrmMain.PicMap.Refresh
End Sub
