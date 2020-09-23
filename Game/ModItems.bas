Attribute VB_Name = "ModItems"
Sub EmptyItem(Index As Integer)
Item(Index).Ammount = 0
Item(Index).MAP = 0
Item(Index).Name = ""
Item(Index).Sprite = 0
Item(Index).Visible = 0
Item(Index).X = 0
Item(Index).Y = 0
End Sub


Sub EmptyINVSlot(Index As Integer)
INV(Index).Ammount = 0
INV(Index).Index = 0
INV(Index).Name = ""
INV(Index).Sprite = 0
End Sub

