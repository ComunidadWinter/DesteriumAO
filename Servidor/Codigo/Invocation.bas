Attribute VB_Name = "Invocation"
Option Explicit

' INVOCACIONES CON USUARIOS

Public Type tInvocaciones
    
    Activo As Byte
    
    
    'INFORMACION CARGADA
    desc As String
    NpcIndex As Integer
    CantidadUsuarios As Byte
    Mapa As Byte
    X() As Byte
    Y() As Byte
    
        
    
    
End Type


Public NumInvocaciones As Byte
Public Invocaciones() As tInvocaciones

'[INIT]
'NumInvocaciones = 1

'[INVOCACION1] 'Mago del inframundo
'NpcIndex = 410

'Mapa = 1
'CantidadUsuarios = 2
'Pos1 = 40 - 60
'Pos2 = 70 - 80
Public Sub LoadInvocaciones()
    
    Dim i As Integer
    Dim X As Integer
    Dim ln As String
    
    NumInvocaciones = val(GetVar(DatPath & "Invocaciones.dat", "INIT", "NumInvocaciones"))
    
    ReDim Invocaciones(1 To NumInvocaciones) As tInvocaciones
        For i = 1 To NumInvocaciones
            With Invocaciones(i)
                .Activo = 0
                .CantidadUsuarios = val(GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "CantidadUsuarios"))
                .Mapa = val(GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "Mapa"))
                .NpcIndex = val(GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "NpcIndex"))
                .desc = GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "Desc")
                
                ReDim .X(1 To .CantidadUsuarios)
                ReDim .Y(1 To .CantidadUsuarios)
                
                
                For X = 1 To .CantidadUsuarios
                    ln = GetVar(DatPath & "Invocaciones.dat", "INVOCACION" & i, "Pos" & X)
                    
                    .X(X) = val(ReadField(1, ln, 45))
                    .Y(X) = val(ReadField(2, ln, 45))
                Next X
            End With
        Next i

    
End Sub

Public Function InvocacionIndex(ByVal Mapa As Byte, ByVal X As Byte, ByVal Y As Byte) As Byte

    Dim i As Integer
    Dim Z As Integer
    
    InvocacionIndex = 0
    
    '// Devuelve el Index del mapa de invocación en el que está
    For i = 1 To NumInvocaciones
        With Invocaciones(i)
            For Z = 1 To .CantidadUsuarios
                If .Mapa = Mapa And (.X(Z) = X) And .Y(Z) = Y Then
                    InvocacionIndex = i
                    Exit For
                End If
            Next Z
        End With
    Next i
            
        
End Function

' if invocacacionindex = 0 then
Public Function PuedeSpawn(ByVal index As Byte) As Boolean
    
    Dim Contador As Byte
    Dim i As Integer
    
    PuedeSpawn = False
    For i = 1 To Invocaciones(index).CantidadUsuarios
        If MapData(Invocaciones(index).Mapa, Invocaciones(index).X(i), Invocaciones(index).Y(i)).UserIndex Then
            Contador = Contador + 1
            
            If Contador = Invocaciones(index).CantidadUsuarios Then
                PuedeSpawn = True
            End If
        End If
    Next i
    
End Function

Public Function PuedeRealizarInvocacion(ByVal UserIndex As Integer) As Boolean
    PuedeRealizarInvocacion = False
    
    With UserList(UserIndex)
        If .flags.Muerto Then Exit Function
        If .flags.Mimetizado Then Exit Function
        
    End With
    
    
    PuedeRealizarInvocacion = True
End Function

Public Sub RealizarInvocacion(ByVal UserIndex As Integer, ByVal index As Byte)
    
    Dim Pos As WorldPos
    
    ' ¿Los usuarios están en las pos?
    If PuedeSpawn(index) Then
        
        Dim NpcIndex As Integer
        Pos.Map = Invocaciones(index).Mapa
        Pos.X = RandomNumber(Invocaciones(index).X(1) - 5, Invocaciones(index).X(1) + 5)
        Pos.Y = RandomNumber(Invocaciones(index).Y(1) - 5, Invocaciones(index).Y(1) + 5)
        
        FindLegalPos UserIndex, Pos.Map, Pos.X, Pos.Y
        NpcIndex = SpawnNpc(Invocaciones(index).NpcIndex, Pos, True, False)
        
        If Not NpcIndex = 0 Then
            Invocaciones(index).Activo = 1
            Npclist(NpcIndex).flags.Invocacion = 1
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Invocaciones(index).desc, FontTypeNames.FONTTYPE_GUILD))
        End If
    End If
    
    
End Sub


