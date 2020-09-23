Attribute VB_Name = "ModDec"
'Just Declairs Bitblt
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'This is out map 200x200 tiles
'There is 1 glitch with this SPEED and this is nothing to do with my codeing
'MAJOR lag can occure if the array is 2 high e.g 9000x9000 map lags on startup
'of the program / Closeing of the program a 1 ghz computer can handle
'1000 x1000 Sooo 2 ghz = 2000 3 ghz = 3000 (this is just an estimate though)
'NOTE: the bigger this array is the bigger the map files when saved will be
'100x100 = 79kb per map!
'200x200 = 158kb
'400x400 = 316 and so on
Public Map(-10 To 210, -10 To 210) As Tile
Public MAPOPEN As Integer
'Just Npc Stuff
Public Const MAX_NPCS = 20
Public Npc(0 To MAX_NPCS) As NPCTYPE

'Basically the same as NPC's but used for items
Public Const MAX_ITEMS = 20
Public Item(0 To MAX_ITEMS) As ITEMTYPE

'Inventory
Public Const MAX_INV = 10
Public INV(0 To MAX_INV) As INVTYPE


'This type is just saying where tthe tile is set
'when you select a tile it gets set in these variabls
Type Tile
WarpMap As Integer
WarpX As Integer
WarpY As Integer
X As Integer
Y As Integer
Blocked As Integer
End Type

'these are used for the NPC INI
Type NPCTYPE
Name As String
Sprite As Integer
Direction As Integer
Map As Integer
X As Integer
Y As Integer
MSG As String
Visible As Integer
Damage As Integer
Range As Integer ' This is how many Tiles away you are b4 the monster attacks
TakeItem As Integer
GiveItem As Integer
End Type

'This is the item type it is just variabls for the items.
Type ITEMTYPE
Name As String
Sprite As Integer
Map As Integer
X As Integer
Y As Integer
Ammount As Integer
Visible As Integer
End Type

'This is your inventory types
Type INVTYPE
Name As String
Ammount As Integer
Sprite As Integer
Index As Integer
End Type

'The offsets are the top left tile where your located (for map scrolling)
Public OffsetX As Integer
Public OffsetY As Integer

'0 = up 1 = down ect
Public PlayerDir As Integer

'Thse are used just to keep tabs on where you at :)
Public PlayerX As Integer
Public PlayerY As Integer

'This is a basic Helth Var
Public PlayerHelth As Integer

'This is a basic Helth Var
Public PlayerStam As Integer

'Players Name
Public PlayerName As String


'HALT MOVEMENT
Public HALTmov As Boolean

'This is to display Attacks
Public Hurting As Boolean
Public HurtX As Integer
Public HurtY As Integer
Public HitAmmount As Integer
