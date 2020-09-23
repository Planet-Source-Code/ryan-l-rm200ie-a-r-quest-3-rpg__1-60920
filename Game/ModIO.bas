Attribute VB_Name = "ModIO"
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long


Sub OPENMAP(Num As Integer)
Open App.Path & "/Maps/" & Num & ".map" For Binary As #1
Get #1, , Map
Close #1

MAPOPEN = Num
End Sub

'Im not going to explain below here much as this is just basic INI reading if u dont know this then you shoudlt be making games.
Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = ""
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub

Function FileExist(ByVal FileName As String) As Boolean
    If Dir(App.Path & "\" & FileName) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function



'LOAD NPCS
Sub LoadNPCs()
Dim FileName As String
Dim i As Long
    
    FileName = App.Path & "\NPC.ini"
    
    For i = 0 To MAX_NPCS
    
    'These open the INI into the varible
        Npc(i).Name = GetVar(FileName, "NPC" & i, "Name")
        Npc(i).Sprite = Val(GetVar(FileName, "NPC" & i, "Sprite"))
        Npc(i).Direction = Val(GetVar(FileName, "NPC" & i, "Direction"))
        Npc(i).Map = Val(GetVar(FileName, "NPC" & i, "MAP"))
        Npc(i).X = Val(GetVar(FileName, "NPC" & i, "X"))
        Npc(i).Y = Val(GetVar(FileName, "NPC" & i, "Y"))
        Npc(i).MSG = GetVar(FileName, "NPC" & i, "MSG")
        Npc(i).Visible = Val(GetVar(FileName, "NPC" & i, "Visible"))
        Npc(i).Damage = Val(GetVar(FileName, "NPC" & i, "Damage"))
        Npc(i).Range = Val(GetVar(FileName, "NPC" & i, "Range"))
        
        Npc(i).TakeItem = Val(GetVar(FileName, "NPC" & i, "TakeItem"))
        Npc(i).GiveItem = Val(GetVar(FileName, "NPC" & i, "GiveItem"))
    'This makes it Blocked where the npc is
    
        If Npc(i).Map = MAPOPEN Then
        Map(Npc(i).X, Npc(i).Y).Blocked = 1
        End If
        
        DoEvents
    Next i
    
Refresh_Screen
End Sub

'LOAD ITEMS
Sub LoadITEMs()
Dim FileName As String
Dim i As Long
    
    FileName = App.Path & "\ITEMS.ini"
    
    For i = 0 To MAX_ITEMS
    
        Item(i).Name = GetVar(FileName, "ITEM" & i, "Name")
        Item(i).Sprite = Val(GetVar(FileName, "ITEM" & i, "Sprite"))
        Item(i).Map = Val(GetVar(FileName, "ITEM" & i, "Map"))
        Item(i).X = Val(GetVar(FileName, "ITEM" & i, "X"))
        Item(i).Y = Val(GetVar(FileName, "ITEM" & i, "Y"))
        Item(i).Ammount = 1 'Reason for this is beacuse the item system isnot yet setup to handle more than 1 item :P'Val(GetVar(FileName, "ITEM" & i, "Ammount"))
        Item(i).Visible = Val(GetVar(FileName, "ITEM" & i, "Visible"))
        
        DoEvents
    Next i
    
Refresh_Screen
End Sub


Sub OPENSTUFF()
LoadNPCs
LoadITEMs

End Sub
