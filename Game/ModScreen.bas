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

'BEGIN ITEM
Call BlitITEMS(X(1), X(2), Y(1), Y(2))
'END ITEM

'BEGIN NPC
Call BlitNPCS(X(1), X(2), Y(1), Y(2))
'END NPC

'These are just for scrolling its like a Loop without the For
'This line Scrolls to the bottem of the screen
Y(2) = Y(2) + 1
Next Y(1)
'Then it gets set to 0 (back to the top)
Y(2) = 0
'Then It moves over 1 tile and starts again
X(2) = X(2) + 1
Next X(1)

'Blits the char in the center of the screen with all his text
FrmMain.PicMap.Font.Size = 10
FrmMain.PicMap.FillColor = vbWhite
FrmMain.PicMap.CurrentX = (5 * 32)
FrmMain.PicMap.CurrentY = (5 * 32) - 10
FrmMain.PicMap.Print PlayerName
BitBlt FrmMain.PicMap.hDC, 5 * 32, 5 * 32, 32, 32, FrmMain.picCharMSK.hDC, PlayerDir * 32, 0, vbSrcPaint
BitBlt FrmMain.PicMap.hDC, 5 * 32, 5 * 32, 32, 32, FrmMain.picChar.hDC, PlayerDir * 32, 0, vbSrcAnd

'If the player is Hurt then Display how much he was hit for then enable the hurt off timer
If Hurting = True Then
BitBlt FrmMain.PicMap.hDC, 5 * 32, 5 * 32, 32, 32, FrmMain.AttMSK.hDC, 0, 0, vbSrcPaint
BitBlt FrmMain.PicMap.hDC, 5 * 32, 5 * 32, 32, 32, FrmMain.Att.hDC, 0, 0, vbSrcAnd
FrmMain.PicMap.Font.Size = 12
FrmMain.PicMap.FillColor = vbYellow
If Len(HitAmmount) = 2 Then
    FrmMain.PicMap.CurrentX = (5 * 32) + 0
    FrmMain.PicMap.CurrentY = (5 * 32) + 5
Else
    FrmMain.PicMap.CurrentX = (5 * 32) + 5
    FrmMain.PicMap.CurrentY = (5 * 32) + 5
End If
FrmMain.PicMap.Print HitAmmount
FrmMain.tmrHurt.Enabled = True
End If

'Adds to the fps
FPS = FPS + 1

'Just says where the player is he is at offsetx + 5 and same for y
PlayerX = OffsetX + 5
PlayerY = OffsetY + 5

'Refresh the main screen without this you wont see anything on the display
FrmMain.PicMap.Refresh
End Sub



Sub BlitNPCS(x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer)
Dim X(1 To 2) As Integer
Dim Y(1 To 2) As Integer
X(1) = x1
X(2) = x2
Y(1) = y1
Y(2) = y2

For s = 0 To MAX_NPCS
If Npc(s).X > PlayerX - 5 And Npc(s).X < PlayerX + 5 And Npc(s).Y > PlayerY - 5 And Npc(s).Y < PlayerY + 5 Then
    If Npc(s).X = X(1) And Npc(s).Y = Y(1) Then
        If Npc(s).Visible = 1 Then
            If Npc(s).Map = MAPOPEN Then
                Call BitBlt(FrmMain.PicMap.hDC, X(2) * 32, Y(2) * 32, 32, 32, FrmMain.picNpcMSK.hDC, Npc(s).Direction * 32, 0, vbSrcPaint)
                Call BitBlt(FrmMain.PicMap.hDC, X(2) * 32, Y(2) * 32, 32, 32, FrmMain.picNPC.hDC, Npc(s).Direction * 32, Npc(s).Sprite * 32, vbSrcAnd)
            End If
        End If
    End If
End If
Next s
End Sub

Sub BlitITEMS(x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer)
Dim X(1 To 2) As Integer
Dim Y(1 To 2) As Integer
X(1) = x1
X(2) = x2
Y(1) = y1
Y(2) = y2

For s = 0 To MAX_ITEMS
If Item(s).X > PlayerX - 5 And Item(s).X < PlayerX + 5 And Item(s).Y > PlayerY - 5 And Item(s).Y < PlayerY + 5 Then
If Item(s).X = X(1) And Item(s).Y = Y(1) Then
If Item(s).Visible = 1 Then
If Item(s).Map = MAPOPEN Then
BitBlt FrmMain.PicMap.hDC, X(2) * 32, Y(2) * 32, 32, 32, FrmMain.picItemsMSK.hDC, 0, Item(s).Sprite * 32, vbSrcPaint
BitBlt FrmMain.PicMap.hDC, X(2) * 32, Y(2) * 32, 32, 32, FrmMain.picItems.hDC, 0, Item(s).Sprite * 32, vbSrcAnd
End If
End If
End If
End If
Next s
End Sub
