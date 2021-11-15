Attribute VB_Name = "NPC_Canje"
Option Explicit
Public Const NpcCanjes As Integer = 1055 'CAMBIEN POR SU NPC
Public NumCanjes As Byte
Public tCanje() As tCanjes
Public CanjesPath As String
 
Public Type tCanjes
  GrhIndex As Integer  'grhindex del obj jeje
  PointsR As Byte      'puntos requeridos
  ObjIndex As Integer  'namber of obj.dat
  num As Byte          'cantidad de obj qe se le da
End Type
 
 
Public Sub Canjes_Load()
  Dim i As Long
  CanjesPath = App.Path & "\Dat\Canjes.txt"
  NumCanjes = val(GetVar(CanjesPath, "INICIO", "NumCanjes"))
  ReDim tCanje(1 To NumCanjes)
  For i = 1 To NumCanjes
  With tCanje(i)
  .GrhIndex = val(GetVar(CanjesPath, "CANJE" & i, "GrhIndex"))
  .PointsR = val(GetVar(CanjesPath, "CANJE" & i, "Puntos"))
  .ObjIndex = val(GetVar(CanjesPath, "CANJE" & i, "Objeto"))
  .num = val(GetVar(CanjesPath, "CANJE" & i, "Num"))
  End With
  Next i
End Sub
 
Public Sub Canjes_AdminDaPoints(ByVal UserIndex As Integer, ByVal tUserI As Integer, ByVal uPoints As Byte)
With UserList(tUserI)
WriteConsoleMsg tUserI, UserList(UserIndex).Name & " te otorgó " & uPoints & " puntos de canje.", FontTypeNames.FONTTYPE_GUILD
.Stats.Points = .Stats.Points + uPoints
WriteUpdateUserStats tUserI
End With
End Sub
 
Public Sub Canjes_uRequiereForm(ByVal UserIndex As Integer)
If UserList(UserIndex).Stats.Points > 0 Then
WriteSendCanjes UserIndex
End If
End Sub
 
Public Sub Canjes_uCanjea(ByVal UserIndex As Integer, ByVal cSlot As Byte)
With UserList(UserIndex)
If Not Canje_Valido(cSlot) Then Exit Sub
 
If tCanje(cSlot).PointsR > .Stats.Points Then
WriteConsoleMsg UserIndex, "Te faltan " & (tCanje(cSlot).PointsR - .Stats.Points) & " puntos de canje para canjear este objeto", FontTypeNames.FONTTYPE_GUILD
Exit Sub
End If
 
Dim tObj As Obj
tObj.Amount = tCanje(cSlot).num
tObj.ObjIndex = tCanje(cSlot).ObjIndex
 
If Not MeterItemEnInventario(UserIndex, tObj) Then Call TirarItemAlPiso(.Pos, tObj)
.Stats.Points = .Stats.Points - tCanje(cSlot).PointsR
WriteConsoleMsg UserIndex, "Has canjeado " & ObjData(tCanje(cSlot).ObjIndex).Name & ". " & IIf(.Stats.Points > 0, "te quedan " & .Stats.Points, "."), FontTypeNames.FONTTYPE_GUILD
 
WriteUpdateUserStats UserIndex
 
End With
End Sub
 
Public Function Canje_Valido(ByVal CanSlot As Byte) As Boolean
If CanSlot < 0 Or CanSlot > NumCanjes Then Canje_Valido = False
Canje_Valido = True
End Function
