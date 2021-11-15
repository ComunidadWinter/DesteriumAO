Attribute VB_Name = "mGranPoder"
' Módulo de gran poder. Idea basada en el viejo código de TPAO pero con el toque de ZOOM
' Si lees esto te violaron de chiquito

Option Explicit

Private Type GreatPower
    LastUser As String
    CurrentUser As String
    CurrentMap As Integer
End Type

Public GreatPower As GreatPower

Public Function UserIndex_GreatPower() As Boolean
    ' • Chequeo el usuario que va a recibir el poder
    ' • Otorgamos el Gran Poder a un usuario Random
    ' • Comentarios por si otro programador toca esto(?)
    
    Dim LoopC As Integer
    Dim UserIndex As Integer
    Dim Exist As Boolean: Exist = False

    If LastUser = 0 Then Exit Function
    
    Do While (Exist = False) And LoopC < 20
        LoopC = LoopC + 1
        UserIndex = RandomNumber(1, LastUser)
         
        With UserList(UserIndex)
            If (.flags.UserLogged = True) And (.flags.Muerto = 0) And (.flags.Privilegios = User) And (.Pos.map <> 176 And .Pos.map <> 191) Then
                If (StrComp(GreatPower.LastUser, UCase$(.Name)) <> 0) And (StrComp(GreatPower.CurrentUser, UCase$(.Name)) <> 0) And _
                    (MapInfo(.Pos.map).Pk = True) Then
                    
                    GreatPower.LastUser = UCase$(GreatPower.CurrentUser)
                    GreatPower.CurrentUser = UCase$(.Name)
                    GreatPower.CurrentMap = .Pos.map
                    UserIndex_GreatPower = True
                    Exist = True
                    Exit Do
                End If
            End If
        End With
    Loop
    
    If UserIndex_GreatPower Then
        With UserList(UserIndex)
            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg( _
                "Gran Poder de los Dioses» Los dioses le han otorgado el poder al personaje " & GreatPower.CurrentUser & _
                " ubicado en el mapa " & GreatPower.CurrentMap & "(" & MapInfo(GreatPower.CurrentMap).Name & ")", FontTypeNames.FONTTYPE_PREMIUM)
            
            RefreshCharStatus UserIndex
        End With
    End If
End Function

Public Function Check_GreatPower(ByVal UserIndex As Integer, _
                                Optional ByVal AttackerIndex As Integer = 0) As Boolean
    Check_GreatPower = True
    
    With UserList(UserIndex)

        
        Exit Function
        
        ' ¿Se fue a zona segura?
        If Not MapInfo(.Pos.map).Pk Then Check_GreatPower = False
        
        ' ¿Deslogea?
        If Not .flags.UserLogged Then Check_GreatPower = False
        
        ' ¿Muerto?
        If .flags.Muerto Then Check_GreatPower = False

        ' USUARIO SIGUE CON GRAN PODER PERO CAMBIO DE MAPA
        If Check_GreatPower Then
            If .Pos.map <> GreatPower.CurrentMap Then
                
                GreatPower.CurrentMap = .Pos.map
                
                If RandomNumber(1, 10) <= 4 Then
                    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg( _
                    "Gran Poder de los Dioses» " & GreatPower.CurrentUser & _
                    " ubicado en el mapa " & GreatPower.CurrentMap & "(" & MapInfo(GreatPower.CurrentMap).Name & ")", FontTypeNames.FONTTYPE_PREMIUM)
                End If
            End If
        End If
        
        ' Se busca nuevo usuario si se pierde por causa "natural"
        If (Check_GreatPower = False) And (AttackerIndex = 0) Then
            GreatPower.LastUser = UCase$(.Name)
            GreatPower.CurrentUser = vbNullString
            GreatPower.CurrentMap = 0
            UserIndex_GreatPower
            
            RefreshCharStatus UserIndex
        ' Muere por un user mas polenta
        ElseIf (Check_GreatPower = False) And (AttackerIndex > 0) Then
            GreatPower.LastUser = UCase$(.Name)
            GreatPower.CurrentUser = UCase$(UserList(AttackerIndex).Name)
            GreatPower.CurrentMap = UserList(AttackerIndex).Pos.map
            RefreshCharStatus UserIndex
            RefreshCharStatus AttackerIndex
            
            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg( _
                "Gran Poder de los Dioses» El poder ha pasado a manos de " & UserList(AttackerIndex).Name & _
                " ubicado en el mapa " & UserList(AttackerIndex).Pos.map & "(" & MapInfo(UserList(AttackerIndex).Pos.map).Name & ")", FontTypeNames.FONTTYPE_PREMIUM)
        End If
    End With
End Function
