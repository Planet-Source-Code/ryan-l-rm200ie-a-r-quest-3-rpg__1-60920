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


'The offsets are the top left tile where your located (for map scrolling)
Public OffsetX As Integer
Public OffsetY As Integer

'This is for the frame per second count
Public FPS As Integer
