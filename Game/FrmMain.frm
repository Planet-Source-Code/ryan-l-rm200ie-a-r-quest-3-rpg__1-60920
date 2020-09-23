VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "R Quest 3 Engine"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7995
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox AttMSK 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   6000
      Picture         =   "FrmMain.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   30
      Top             =   960
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox Att 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   5400
      Picture         =   "FrmMain.frx":0C42
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   29
      Top             =   960
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      ScaleHeight     =   225
      ScaleWidth      =   2385
      TabIndex        =   28
      Top             =   480
      Width           =   2415
      Begin VB.Shape shpStam 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      ScaleHeight     =   225
      ScaleWidth      =   2385
      TabIndex        =   27
      Top             =   120
      Width           =   2415
      Begin VB.Shape shpHelth 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.PictureBox picEmpty 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   10560
      Picture         =   "FrmMain.frx":1884
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   24
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picCharMSK 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   8400
      Picture         =   "FrmMain.frx":24C6
      ScaleHeight     =   480
      ScaleWidth      =   1920
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.PictureBox picItemsMSK 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   5880
      Picture         =   "FrmMain.frx":5508
      ScaleHeight     =   1920
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picItems 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   6480
      Picture         =   "FrmMain.frx":854A
      ScaleHeight     =   1920
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Timer tmrTXT 
      Interval        =   2500
      Left            =   4680
      Top             =   4080
   End
   Begin VB.PictureBox picNpcMSK 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9630
      Left            =   8400
      Picture         =   "FrmMain.frx":B58C
      ScaleHeight     =   640
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.PictureBox picNPC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9630
      Left            =   8400
      Picture         =   "FrmMain.frx":475CE
      ScaleHeight     =   640
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   8400
      Picture         =   "FrmMain.frx":83610
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.PictureBox PicTiles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9150
      Left            =   11160
      Picture         =   "FrmMain.frx":86652
      ScaleHeight     =   608
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   288
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   4350
   End
   Begin VB.PictureBox PicMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   5295
      Left            =   0
      ScaleHeight     =   351
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   351
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.Timer tmrHurt 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   4680
         Top             =   2880
      End
      Begin VB.Timer tmrStam 
         Interval        =   2000
         Left            =   4680
         Top             =   3480
      End
      Begin VB.PictureBox PicInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   960
         ScaleHeight     =   1785
         ScaleWidth      =   3465
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   3495
         Begin VB.PictureBox invItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   10
            Left            =   2400
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   23
            Top             =   720
            Width           =   495
         End
         Begin VB.PictureBox invItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   9
            Left            =   1920
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   22
            Top             =   720
            Width           =   495
         End
         Begin VB.PictureBox invItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   8
            Left            =   1440
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   21
            Top             =   720
            Width           =   495
         End
         Begin VB.PictureBox invItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   7
            Left            =   960
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   20
            Top             =   720
            Width           =   495
         End
         Begin VB.PictureBox invItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   6
            Left            =   480
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   19
            Top             =   720
            Width           =   495
         End
         Begin VB.PictureBox invItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   5
            Left            =   2640
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   18
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox invItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   4
            Left            =   2160
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   17
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox invItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   3
            Left            =   1680
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   16
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox invItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   2
            Left            =   1200
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   15
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox invItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   1
            Left            =   720
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   14
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox invItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   0
            Left            =   240
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H8000000C&
            Caption         =   "Inventory"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   26
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label lblSH 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Close (X)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2160
            TabIndex        =   12
            Top             =   1320
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.Timer tmrFollow 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   4680
         Top             =   4680
      End
      Begin VB.Label lblMSG 
         BackStyle       =   0  'Transparent
         Caption         =   "Messages popup in this label."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   4080
         Visible         =   0   'False
         Width           =   5055
      End
   End
   Begin VB.Label Label2 
      Caption         =   "5"
      Height          =   255
      Left            =   8400
      TabIndex        =   4
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "5"
      Height          =   255
      Left            =   8400
      TabIndex        =   3
      Top             =   3120
      Width           =   2055
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim LastNPC As Integer


Private Sub Form_Load()
OPENMAP 1
OPENSTUFF
PlayerHelth = 100
PlayerStam = 100
PlayerName = "Test"
End Sub


Private Sub invItem_Click(Index As Integer)
Label3.Caption = INV(Index).Name
End Sub

Private Sub invItem_DblClick(Index As Integer)
If Not Len(INV(Index).Name) > 0 Then Exit Sub

Item(INV(Index).Index).Name = INV(Index).Name
Item(INV(Index).Index).Ammount = INV(Index).Ammount
Item(INV(Index).Index).Map = MAPOPEN
Item(INV(Index).Index).X = PlayerX
Item(INV(Index).Index).Y = PlayerY
Item(INV(Index).Index).Sprite = INV(Index).Sprite
If Len(INV(Index).Name) > 1 Then
Item(INV(Index).Index).Visible = 1
End If
EmptyINVSlot (Index)
Refresh_Screen
'lblSH_Click

For A = 0 To MAX_INV
If Not INV(A).Name = "" Then
    BitBlt invItem(A).hDC, 0, 0, 32, 32, picItems.hDC, 0, INV(A).Sprite * 32, vbSrcCopy
    invItem(A).Refresh
Else
    BitBlt invItem(A).hDC, 0, 0, 32, 32, picEmpty.hDC, 0, 0, vbSrcCopy
    invItem(A).Refresh
End If
Next A
End Sub

Private Sub lblSH_Click()
PicInv.Visible = False
HALTmov = False
End Sub

Private Sub PicMap_KeyDown(KeyCode As Integer, Shift As Integer)
Dim A As Integer 'just makes it so we dont get type mismatches.
'If you not allowed to move dont take any input
If HALTmov = True Then GoTo Skip:

If KeyCode = vbKeyA Then
tmrMagic.Enabled = True
End If


'Show Inventory
If KeyCode = vbKeyI Then
PicInv.Visible = True
For A = 0 To MAX_INV
If Not INV(A).Name = "" Then
    BitBlt invItem(A).hDC, 0, 0, 32, 32, picItems.hDC, 0, INV(A).Sprite * 32, vbSrcCopy
    invItem(A).Refresh
Else
    BitBlt invItem(A).hDC, 0, 0, 32, 32, picEmpty.hDC, 0, 0, vbSrcCopy
    invItem(A).Refresh
End If
Next A
lblSH.Visible = True
HALTmov = True
End If

'Turn On Move
If KeyCode = vbKeyT Then
If tmrFollow.Enabled = False Then
tmrFollow.Enabled = True
Else
tmrFollow.Enabled = False
LoadNPCs
End If
End If

'Chat to npc's
If KeyCode = vbKeySpace Then
For A = 0 To MAX_NPCS
'NPC's
If Npc(A).X = PlayerX And Npc(A).Y = PlayerY - 1 Then
lblMSG.Caption = Npc(A).MSG
lblMSG.Visible = True
tmrTXT.Enabled = True
HALTmov = True
LastNPC = A
Exit Sub
End If
'ITEM's
If Item(A).X = PlayerX And Item(A).Y = PlayerY Then
For b = 0 To MAX_INV
If INV(b).Name = "" Then
INV(b).Name = Item(A).Name
INV(b).Ammount = Item(A).Ammount
INV(b).Sprite = Item(A).Sprite
INV(b).Index = A
Call EmptyItem(A)
End If
Next b
End If

Next A
End If


'These are pretty self explanitory if u press up Move up lol
If KeyCode = vbKeyUp Then
'Set the players Direction
PlayerDir = 0
'Check Up block
If Map(PlayerX, PlayerY - 1).Blocked = 1 Then GoTo Skip:
If Map(PlayerX, PlayerY - 1).WarpMap > 0 Then Call WarpPlayer(Map(PlayerX, PlayerY - 1).WarpMap, Map(PlayerX, PlayerY - 1).WarpX, Map(PlayerX, PlayerY - 1).WarpY): Exit Sub
'Move
OffsetY = OffsetY - 1
End If

If KeyCode = vbKeyDown Then
'Set the players Direction
PlayerDir = 1
'Check Down Block
If Map(PlayerX, PlayerY + 1).Blocked = 1 Then GoTo Skip:
If Map(PlayerX, PlayerY + 1).WarpMap > 0 Then Call WarpPlayer(Map(PlayerX, PlayerY + 1).WarpMap, Map(PlayerX, PlayerY + 1).WarpX, Map(PlayerX, PlayerY + 1).WarpY): Exit Sub
'Move
OffsetY = OffsetY + 1
End If

If KeyCode = vbKeyLeft Then
'Set the players Direction
PlayerDir = 2
'Check Left Block
If Map(PlayerX - 1, PlayerY).Blocked = 1 Then GoTo Skip:
If Map(PlayerX - 1, PlayerY).WarpMap > 0 Then Call WarpPlayer(Map(PlayerX - 1, PlayerY).WarpMap, Map(PlayerX - 1, PlayerY).WarpX, Map(PlayerX - 1, PlayerY).WarpY): Exit Sub
'Move
OffsetX = OffsetX - 1
End If

If KeyCode = vbKeyRight Then
'Set the players Direction
PlayerDir = 3
'Check Right Block
If Map(PlayerX + 1, PlayerY).Blocked = 1 Then GoTo Skip:
If Map(PlayerX + 1, PlayerY).WarpMap > 0 Then Call WarpPlayer(Map(PlayerX + 1, PlayerY).WarpMap, Map(PlayerX + 1, PlayerY).WarpX, Map(PlayerX + 1, PlayerY).WarpY): Exit Sub
'Move
OffsetX = OffsetX + 1
End If

'because -5 willset player at center (since he is 5 tiles in the middle)
If OffsetX < -5 Then OffsetX = -5
If OffsetY < -5 Then OffsetY = -5
'Same as above but this is just for the map boundry on the other side
If OffsetX > 195 Then OffsetX = 195
If OffsetY > 195 Then OffsetY = 195

'Redraw the screen
Refresh_Screen
Label1.Caption = PlayerX
Label2.Caption = PlayerY
PlayerStam = PlayerStam - 1
RenderStat

Skip:
Refresh_Screen
End Sub


Private Sub tmrFollow_Timer()
Dim A As Integer
For A = 0 To MAX_NPCS
If Npc(A).Damage > 0 Then
Call procMoveNPC(A)
End If
Next A
End Sub


Private Sub tmrHurt_Timer()
Hurting = False
Refresh_Screen
tmrHurt.Enabled = False
End Sub

Private Sub tmrStam_Timer()
PlayerStam = PlayerStam + 5
PlayerHelth = PlayerHelth + 5
RenderStat
End Sub

Private Sub tmrTXT_Timer()

If Npc(LastNPC).TakeItem > 0 Then
    For A = 0 To MAX_INV
        If Item(Npc(LastNPC).TakeItem).Name = INV(A).Name Then
            If MsgBox("You are holding the item this npc wants will you give " & Item(Npc(LastNPC).TakeItem).Name & " for a " & Item(Npc(LastNPC).GiveItem).Name, vbYesNo, "Quest!") = vbYes Then
                EmptyINVSlot (A)
                INV(b).Name = Item(Npc(LastNPC).GiveItem).Name
                INV(b).Ammount = Item(Npc(LastNPC).GiveItem).Ammount
                INV(b).Sprite = Item(Npc(LastNPC).GiveItem).Sprite
                INV(b).Index = Npc(LastNPC).GiveItem
            End If
        End If
    Next A
End If

lblMSG.Visible = False
HALTmov = False
tmrTXT.Enabled = False
End Sub
