Attribute VB_Name = "ModNPC"
'Monsters monsters monsters!!! im not very good with these i couldnt really
'work out how they work properlly but ive got some basic BASIC!!! stuff down
'and its free for all you out there to use if you want but if you update the monster
'code i really would love an email with the update for the monsters please
'rm2kdev@wasp.net.au

Public Sub procMoveNPC(Monno As Integer)
'Get distance to hero
d2h = Abs(Npc(Monno).X - PlayerX)

'If d2h > 3 Then
'DX = 1 - Int(Rnd * 3) ' random movement
'DY = 1 - Int(Rnd * 3)
'End If

'if your inrange then move the npc
If d2h <= Npc(Monno).Range Then
DX = 0

If Npc(Monno).X > PlayerX Then
If Map(Npc(Monno).X - 1, Npc(Monno).Y).Blocked = 1 Then Exit Sub
DX = -1
End If

If Npc(Monno).X < PlayerX Then
If Map(Npc(Monno).X + 1, Npc(Monno).Y).Blocked = 1 Then Exit Sub
DX = 1
End If

DY = 0
If Npc(Monno).Y > PlayerY Then
If Map(Npc(Monno).X, Npc(Monno).Y - 1).Blocked = 1 Then Exit Sub
DY = -1
End If

If Npc(Monno).Y < PlayerY Then
If Map(Npc(Monno).X, Npc(Monno).Y + 1).Blocked = 1 Then Exit Sub
DY = 1
End If
End If

'Update the values
Npc(Monno).X = Npc(Monno).X + DX
Npc(Monno).Y = Npc(Monno).Y + DY

If Npc(Monno).X = PlayerX And Npc(Monno).Y = PlayerY Then
HitFOR = Int(Int(Rnd * 3) + Npc(Monno).Damage / Npc(nonno).Range)
PlayerHelth = PlayerHelth - HitFOR
Hurting = True
HurtX = PlayerX - 5
HurtY = PlayerX - 5
HitAmmount = HitFOR
End If

If PlayerHelth <= 0 Then
MsgBox "You Have Died"
End
'Put yout own Death Sub Here as you can see :D mines not very Good (^.^)
End If

Refresh_Screen
RenderStat
End Sub
