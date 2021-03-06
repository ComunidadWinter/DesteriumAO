Attribute VB_Name = "Mod_TileEngine"

Option Explicit

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

Private Const GrhFogata As Integer = 1521

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1


'Encabezado bmp
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    X As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    map As Integer
    X As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tama?o y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    Speed As Single
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type


'Apariencia del personaje
Public Type Char
    MinHp As Integer
    MaxHp As Integer
    MinMan As Integer
    MaxMan As Integer
    Movimient As Boolean
    LastDialog     As String
    Active As Byte
    Heading As E_Heading
    Pos As Position
    
        cCont As Integer
   Drawers As Integer
   
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    
    fX As Grh
    FxIndex As Integer
    
    Team As Byte
    
    Criminal As Byte
    Atacable As Boolean
    
    Nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    Infected As Byte
    Angel As Byte
    Demonio As Byte
    
    priv As Byte
End Type

'Info de un objeto
Public Type Obj
    ObjIndex As Integer
    amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Damage As DList
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
End Type

'DX7 Objects
Public DirectX As New DirectX7
Public DirectDraw As DirectDraw7
Private PrimarySurface As DirectDrawSurface7
Private PrimaryClipper As DirectDrawClipper
Private BackBufferSurface As DirectDrawSurface7

Public IniPath As String
Public MapPath As String


'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long
Private fpsLastCheck As Long

'Tama?o del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer

'Offset del desde 0,0 del main view
Private MainViewTop As Integer
Private MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tama?o muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

'Tama?o de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Dim timerElapsedTime As Single
Dim timerTicksPerFrame As Single
Dim engineBaseSpeed As Single


Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer


Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Private MainViewWidth As Integer
Private MainViewHeight As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte




'???????????????????Graficos??????????????????????
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
'??????????????????????????????????????????????????

'???????????????????Mapa???????????????????????
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'??????????????????????????????????????????????????

Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long

Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer

Public charlist(1 To 10000) As Char

#If SeguridadAlkon Then

Public mi(1 To 1233) As clsManagerInvisibles
Public CualMI As Integer

#End If

' Used by GetTextExtentPoint32
Private Type Size
    cx As Long
    cy As Long
End Type

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'??????????????????????????????????????????????????

'#If ConAlfaB Then

Private Declare Function BltAlphaFast Lib "vbabdx" (ByRef lpDDSDest As Any, ByRef lpDDSSource As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchSrc As Long, ByVal pitchDst As Long, ByVal dwMode As Long) As Long
Private Declare Function BltEfectoNoche Lib "vbabdx" (ByRef lpDDSDest As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchDst As Long, ByVal dwMode As Long) As Long

'#End If
'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'Text width computation. Needed to center text.
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Sub CargarCabezas()
    Dim n As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    n = FreeFile()
    Open App.path & "\Recursos\Clases\Clases.ind" For Binary Access Read As #n
    
    If Not FileExist(App.path & "\INIT\CABEZAS.IND", vbArchive) Then
End
End If

If Not FileExist(App.path & "\Recursos\Clases\Clases.ind", vbArchive) Then
End
End If
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #n, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #n
End Sub

Sub CargarCascos()
    Dim n As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    n = FreeFile()
    Open App.path & "\init\Cascos.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #n, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #n
End Sub

Sub CargarCuerpos()
    Dim n As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    n = FreeFile()
    Open App.path & "\init\Personajes.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #n, , MisCuerpos(i)
        
            If MisCuerpos(i).Body(1) Then
                InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
                InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
                InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
                InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
                
                BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
                BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
            End If
    Next i
    
    Close #n
End Sub

Sub CargarFxs()
    Dim n As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    n = FreeFile()
    Open App.path & "\init\Fxs.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #n, , FxData(i)
    Next i
    
    Close #n
End Sub

Sub CargarTips()
    Dim n As Integer
    Dim i As Long
    Dim NumTips As Integer
    
    n = FreeFile
    Open App.path & "\init\Tips.ayu" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , NumTips
    
    'Resize array
    ReDim Tips(1 To NumTips) As String * 255
    
    For i = 1 To NumTips
        Get #n, , Tips(i)
    Next i
    
    Close #n
End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .Active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        'Make active
        .Active = 1
    End With
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)

    
    With charlist(CharIndex)
        .Active = 0
        .Team = 0
        .Criminal = 0
        .Atacable = False
        .FxIndex = 0
        .invisible = False
        .Infected = 0
        .Angel = 0
        .Demonio = 0
        
#If SeguridadAlkon Then
        Call mi(CualMI).ResetInvisible(CharIndex)
#End If
        
        .Moving = 0
        .muerto = False
        .Nombre = ""
        .pie = False
        .Pos.X = 0
        .Pos.Y = 0
        .UsandoArma = False
    End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next
    charlist(CharIndex).Active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until charlist(LastChar).Active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y).CharIndex = 0
    
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    
    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.GrhIndex = GrhIndex
    
    If GrhIndex = 0 Then Exit Sub
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addX As Integer
    Dim addY As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addY = -1
        
            Case E_Heading.EAST
                addX = 1
        
            Case E_Heading.SOUTH
                addY = 1
            
            Case E_Heading.WEST
                addX = -1
        End Select
        
        nX = X + addX
        nY = Y + addY
        'Creditos a la puta de miqueas
        If nX < 1 Or nX > 100 Or nY < 1 Or nY > 100 Then Exit Sub
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        MapData(X, Y).CharIndex = 0
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addX
        .scrollDirectionY = addY
    End With
    
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        If CharIndex <> UserCharIndex Then
            Call EraseChar(CharIndex)
        End If
    End If
End Sub

Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", location.X, location.Y, LoopStyle.Enabled)
    End If
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

  Sub DoPasosFx(ByVal CharIndex As Integer)
  Dim paso As Byte
  If UserMontando = True Then Exit Sub
    If Not UserNavegando Then
        With charlist(CharIndex)
            If .muerto = False And EstaPCarea(CharIndex) = True And (.priv <> 5) Then
                .pie = Not .pie
                paso = Map_GetTerrenoDePaso(GrhData(MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex).FileNum)
         
                If paso = 1 Then
                    If .pie Then
                        Call Audio.PlayWave(SND_PASOS7)
                    Else
                        Call Audio.PlayWave(SND_PASOS8)
                    End If
                ElseIf paso = 2 Or paso = 5 Then
                    If .pie Then
                        Call Audio.PlayWave(SND_PASOS1) 'Si no son Pasto ,nieve o arena,se pone este default, 1 y 2!
                    Else
                        Call Audio.PlayWave(SND_PASOS2) 'Si no son Pasto ,nieve o arena,se pone este default, 1 y 2!
                    End If
                ElseIf paso = 3 Then
                    If .pie Then
                        Call Audio.PlayWave(SND_PASOS5)
                    Else
                        Call Audio.PlayWave(SND_PASOS6)
                    End If
                ElseIf paso = 4 Then
                    If .pie Then
                       Call Audio.PlayWave(SND_PASOS3)
                    Else
                       Call Audio.PlayWave(SND_PASOS4)
                    End If
                End If
            End If
        End With
   ' ElseIf UserMontando = True Then
     '   Call Audio.PlayWave(23, charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.y)
    ElseIf UserNavegando = True Then
    '- Saque el sonido del Wather.
    End If
End Sub
Private Function Map_GetTerrenoDePaso(ByVal TerrainFileNum As Integer) As Byte
  If (TerrainFileNum >= 6000 And TerrainFileNum <= 6004) Or (TerrainFileNum >= 550 And TerrainFileNum <= 552) Or (TerrainFileNum >= 6018 And TerrainFileNum <= 6020) Then
  Map_GetTerrenoDePaso = 1
  Exit Function
  ElseIf (TerrainFileNum >= 7501 And TerrainFileNum <= 7507) Or (TerrainFileNum = 7500 Or TerrainFileNum = 7508 Or TerrainFileNum = 1533 Or TerrainFileNum = 2508) Then
  Map_GetTerrenoDePaso = 2
  Exit Function
  ElseIf (TerrainFileNum >= 10139 And TerrainFileNum <= 10143) Then
  Map_GetTerrenoDePaso = 3
  Exit Function
  ElseIf TerrainFileNum = 6021 Then
  Map_GetTerrenoDePaso = 4
  Exit Function
  Else
  Map_GetTerrenoDePaso = 5
  End If
End Function

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim X As Integer
    Dim Y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        MapData(X, Y).CharIndex = 0
        
        addX = nX - X
        addY = nY - Y
        
        If Sgn(addX) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addX) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addY) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addY) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addX)
        .scrollDirectionY = Sgn(addY)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDE Or .FxIndex = FxMeditar.M_PREM Or .FxIndex = FxMeditar.M_ORO Or .FxIndex = FxMeditar.M_ARMI Or .FxIndex = FxMeditar.M_LEGIO Then
            .FxIndex = 0
        End If
        
    End With
    
    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
        
        Case E_Heading.EAST
            X = 1
        
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
    End If
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    location.X = j
                    location.Y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim LoopC As Long
    Dim Dale As Boolean
    
    LoopC = 1
    Do While charlist(LoopC).Active And Dale
        LoopC = LoopC + 1
        Dale = (LoopC <= UBound(charlist))
    Loop
    
    NextOpenChar = LoopC
End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhData() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    
    'Open files
    handle = FreeFile()
    
    Open IniPath & GraphicsFile For Binary Access Read As handle
    Seek #1, 1
    
    'Get file version
    Get handle, , fileVersion
    
    'Get number of grhs
    Get handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(0 To grhCount) As GrhData
    
    While Not EOF(handle)
        Get handle, , Grh
        
        With GrhData(Grh)
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next Frame
                
                Dim sngSpeed As Single
                
                Get handle, , sngSpeed
                
                .Speed = sngSpeed '+ 16.3333

                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , .FileNum
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get handle, , GrhData(Grh).sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = Grh
            End If
        End With
    Wend
    
    Close handle
    
    LoadGrhData = True
Exit Function

ErrorHandler:
    LoadGrhData = False
End Function

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    '?Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        Exit Function
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    bTecho = (MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4)
                       
            If bTecho And UserMontando = True Then Exit Function
    
    LegalPos = True
End Function

Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 01/08/2009
'Checks to see if a tile position is legal, including if there is a casper in the tile
'10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
'01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
'*****************************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(X, Y).CharIndex
    '?Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.X, UserPos.Y).Blocked = 1 Then
            Exit Function
        End If
        
        With charlist(CharIndex)
            ' Si no es casper, no puede pasar
            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.X, UserPos.Y) Then
                    If Not HayAgua(X, Y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(X, Y) Then Exit Function
                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If charlist(UserCharIndex).priv > 0 And charlist(UserCharIndex).priv < 6 Then
                    If charlist(UserCharIndex).invisible = True Or UserOcu = 1 Then Exit Function
                End If
            End If
        End With
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    MoveToLegalPos = True
End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Private Sub DDrawGrhtoSurface(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal center As Byte, ByVal Animate As Byte)
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
On Error GoTo error
        
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * Movement_Speed
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                If Grh.GrhIndex <> 0 Then Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'If Grh.GrhIndex = 0 Then Exit Sub
    
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.GrhIndex <> 0 Then CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_WAIT)
    End With
Exit Sub

error:
    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurri? un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripci?n del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.number & " ] Error"
        End
    End If
End Sub

Sub DDrawTransGrhIndextoSurface(ByVal GrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal center As Byte)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    End With
End Sub

Sub DDrawTransGrhtoSurface(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal center As Byte, ByVal Animate As Byte)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    
On Error GoTo error
    
   If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    End With
Exit Sub

error:
    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurri? un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripci?n del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.number & " ] Error"
        End
    End If
End Sub

Sub DDrawTransGrhtoSurfaceAlpha(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal center As Byte, ByVal Animate As Byte)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    
    If Animate Then
     If ConAlfaB = 1 Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
   End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrentGrhIndex).TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
        End If
        If GrhData(CurrentGrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If
    
    With SourceRect
        .Left = GrhData(CurrentGrhIndex).sX
        .Top = GrhData(CurrentGrhIndex).sY
        .Right = .Left + GrhData(CurrentGrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrhIndex).pixelHeight
    End With
    
    Dim Src As DirectDrawSurface7
    Dim rDest As RECT
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
    
    Set Src = SurfaceDB.Surface(GrhData(CurrentGrhIndex).FileNum)
    
    Src.GetSurfaceDesc ddsdSrc
    BackBufferSurface.GetSurfaceDesc ddsdDest
    
    With rDest
        .Left = X
        .Top = Y
        .Right = X + GrhData(CurrentGrhIndex).pixelWidth
        .Bottom = Y + GrhData(CurrentGrhIndex).pixelHeight
        
        If .Right > ddsdDest.lWidth Then
            .Right = ddsdDest.lWidth
        End If
        If .Bottom > ddsdDest.lHeight Then
            .Bottom = ddsdDest.lHeight
        End If
    End With
    
    ' 0 -> 16 bits 555
    ' 1 -> 16 bits 565
    ' 2 -> 16 bits raro (Sin implementar)
    ' 3 -> 24 bits
    ' 4 -> 32 bits
    
    If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0& And ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0& Then
        Modo = 0
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0& And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0& Then
        Modo = 1
'TODO : Revisar las m?scaras de 24!! Quiz?s mirando el campo lRGBBitCount para diferenciar 24 de 32...
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0& And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0& Then
        Modo = 3
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &HFF00& And ddsdSrc.ddpfPixelFormat.lGBitMask = &HFF00& Then
        Modo = 4
    Else
        'Modo = 2 '16 bits raro ?
        Call BackBufferSurface.BltFast(X, Y, Src, SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        Exit Sub
    End If
    
    Dim SrcLock As Boolean
    Dim DstLock As Boolean
    
    SrcLock = False
    DstLock = False
    
On Local Error GoTo HayErrorAlpha
    
    Call Src.Lock(SourceRect, ddsdSrc, DDLOCK_WAIT, 0)
    SrcLock = True
    Call BackBufferSurface.Lock(rDest, ddsdDest, DDLOCK_WAIT, 0)
    DstLock = True
    
    Call BackBufferSurface.GetLockedArray(dArray())
    Call Src.GetLockedArray(sArray())
    
    Call BltAlphaFast(ByVal VarPtr(dArray(X + X, Y)), ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, Modo)
    
    BackBufferSurface.Unlock rDest
    DstLock = False
    Src.Unlock SourceRect
    SrcLock = False
Exit Sub

HayErrorAlpha:
    If SrcLock Then Src.Unlock SourceRect
    If DstLock Then BackBufferSurface.Unlock rDest
End Sub

Function GetBitmapDimensions(ByVal BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
    Dim BMHeader As BITMAPFILEHEADER
    Dim BINFOHeader As BITMAPINFOHEADER
    
    Open BmpFile For Binary Access Read As #1
    
    Get #1, , BMHeader
    Get #1, , BINFOHeader
    
    Close #1
    
    bmWidth = BINFOHeader.biWidth
    bmHeight = BINFOHeader.biHeight
End Function

Sub DrawGrhtoHdc(ByVal hdc As Long, ByVal GrhIndex As Integer, ByRef SourceRect As RECT, ByRef destRect As RECT)
'*****************************************************************
'Draws a Grh's portion to the given area of any Device Context
'*****************************************************************
    Call SurfaceDB.Surface(GrhData(GrhIndex).FileNum).BltToDC(hdc, SourceRect, destRect)
End Sub

Public Sub DrawTransparentGrhtoHdc(ByVal dsthdc As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal GrhIndex As Integer, ByRef SourceRect As RECT, ByVal TransparentColor As Long)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 12/22/2009
'This method is SLOW... Don't use in a loop if you care about
'speed!
'*************************************************************
    Dim color As Long
    Dim X As Long
    Dim Y As Long
    Dim srchdc As Long
    Dim Surface As DirectDrawSurface7
    
    Set Surface = SurfaceDB.Surface(GrhData(GrhIndex).FileNum)
    
    srchdc = Surface.GetDC
    
    For X = SourceRect.Left To SourceRect.Right - 1
        For Y = SourceRect.Top To SourceRect.Bottom - 1
            color = GetPixel(srchdc, X, Y)
            
            If color <> TransparentColor Then
                Call SetPixel(dsthdc, dstX + (X - SourceRect.Left), dstY + (Y - SourceRect.Top), color)
            End If
        Next Y
    Next X
    
    Call Surface.ReleaseDC(srchdc)
End Sub

Public Sub DrawImageInPicture(ByRef PictureBox As PictureBox, ByRef Picture As StdPicture, ByVal X1 As Single, ByVal Y1 As Single, Optional Width1, Optional Height1, Optional X2, Optional Y2, Optional Width2, Optional Height2)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 12/28/2009
'Draw Picture in the PictureBox
'*************************************************************

Call PictureBox.PaintPicture(Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2)
End Sub



Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Mart?n Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim Y           As Long     'Keeps track of where on map we are
    Dim X           As Long     'Keeps track of where on map we are
    Dim screenminY  As Integer  'Start Y pos on current screen
    Dim screenmaxY  As Integer  'End Y pos on current screen
    Dim screenminX  As Integer  'Start X pos on current screen
    Dim screenmaxX  As Integer  'End X pos on current screen
    Dim minY        As Integer  'Start Y pos on current map
    Dim maxY        As Integer  'End Y pos on current map
    Dim minX        As Integer  'Start X pos on current map
    Dim maxX        As Integer  'End X pos on current map
    Dim ScreenX     As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY     As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset  As Integer
    Dim minYOffset  As Integer
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
     
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            
            'Layer 1 **********************************
            Call DDrawGrhtoSurface(MapData(X, Y).Graphic(1), _
                (ScreenX - 1) * TilePixelWidth + PixelOffsetX + TileBufferPixelOffsetX, _
                (ScreenY - 1) * TilePixelHeight + PixelOffsetY + TileBufferPixelOffsetY, _
                0, 1)
            '******************************************
            
            ScreenX = ScreenX + 1
        Next X
        
        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y
    
    'Draw floor layer 2
    ScreenY = minYOffset
    For Y = minY To maxY
        ScreenX = minXOffset
        For X = minX To maxX
            
            'Layer 2 **********************************
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(2), _
                        (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                        (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                        1, 1)
            End If
            '******************************************
            
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    'Draw Transparent Layers
    ScreenY = minYOffset
    For Y = minY To maxY
        ScreenX = minXOffset
        For X = minX To maxX
            PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffsetY
            
            With MapData(X, Y)
                'Object Layer **********************************
                If .ObjGrh.GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface(.ObjGrh, _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                End If
                '***********************************************
                
                
                'Char layer ************************************
                If .CharIndex <> 0 Then
                    Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                End If
                '*************************************************
                
                'Dibujado del da?o en el Render
                If .Damage.Activated Then
                If TSetup.bGameCombat Then
                   m_Damages.Draw X, Y, PixelOffsetXTemp + 17, PixelOffsetYTemp - -2
                End If
                End If
                
                If ConAlfaB = 0 Then
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    'Draw
                    Call DDrawTransGrhtoSurface(.Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
              End If
            End If
            If ConAlfaB = 1 Then
            
                            If .Graphic(3).GrhIndex = 735 Or .Graphic(3).GrhIndex >= 6994 And .Graphic(3).GrhIndex <= 7002 Then
                If Abs(UserPos.X - X) < 4 And (Abs(UserPos.Y - Y)) < 4 Then
                Call DDrawTransGrhtoSurfaceAlpha(.Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                Else
                Call DDrawTransGrhtoSurface(.Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                End If
                Else
                If .Graphic(3).GrhIndex <> 0 Then
                Call DDrawTransGrhtoSurface(.Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                End If
                End If
                End If
            
                '************************************************
            End With
            
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    
            If ConAlfaB = 1 Then
        'Draw blocked tiles and grid
     ScreenY = minYOffset
        For Y = minY To maxY
            ScreenX = minXOffset
            For X = minX To maxX
 
                'Layer 4 **********************************
             
                If MapData(X, Y).Graphic(4).GrhIndex Then
                    'Draw
                        If alphaT <> 255 Then
                            Call DDrawTransGrhtoSurfaceAlphaTecho(BackBufferSurface, MapData(X, Y).Graphic(4), _
                                 (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                                 (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                                 1, 1, alphaT)
                        Else
                            Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), _
                                 (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                                 (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                                 1, 1)
                        End If
                End If
                '**********************************
 
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
        End If

    If ConAlfaB = 0 Then
    If Not bTecho Then
        'Draw blocked tiles and grid
        ScreenY = minYOffset
        For Y = minY To maxY
            ScreenX = minXOffset
            For X = minX To maxX
                
                'Layer 4 **********************************
                If MapData(X, Y).Graphic(4).GrhIndex Then
                    'Draw
                    Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), _
                        (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                        (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                        1, 1)
                End If
                '**********************************
                
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If
    End If
    
    If Iscombate = True Then
            Call RenderText(260, 260, "Modo Combate", vbRed, frmMain.font)
    End If
    
    If frmMain.macrotrabajo Then
        RenderTextCentered 300, 260, "(Trabajando)", RGB(255, 255, 255), frmMain.font
    End If

    
    Call RenderText(260, 270, "Canjes: " & UserPuntos, vbGreen, frmMain.font)

End Sub

Public Function RenderSounds()
'**************************************************************
'Author: Juan Mart?n Sotuyo Dodero
'Last Modify Date: 3/30/2008
'Actualiza todos los sonidos del mapa.
'**************************************************************
    
    DoFogataFx
End Function

Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean
    If GrhIndex > 0 Then
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
                And charlist(UserCharIndex).Pos.Y <= Y
    End If
End Function

Sub LoadGraphics()
'**************************************************************
'Author: Juan Mart?n Sotuyo Dodero - complete rewrite
'Last Modify Date: 11/03/2006
'Initializes the SurfaceDB and sets up the rain rects
'**************************************************************
    'New surface manager :D
    Call SurfaceDB.Initialize(DirectDraw, ClientSetup.bUseVideo, DirGraficos, ClientSetup.byMemory)
    
    'Set up te rain rects
    RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
    RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
    RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
    RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128
    
    RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
    RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
    RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
    RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
End Sub

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setMainViewTop As Integer, ByVal setMainViewLeft As Integer, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer, ByVal engineSpeed As Single) As Boolean
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Mart?n Sotuyo Dodero (Maraxus)
'Creates all DX objects and configures the engine to start running.
'***************************************************
    Dim SurfaceDesc As DDSURFACEDESC2
    Dim ddck As DDCOLORKEY
    
    IniPath = App.path & "\Init\"
    
    Movement_Speed = 1
    
    'Fill startup variables
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    HalfWindowTileHeight = setWindowTileHeight \ 2
    HalfWindowTileWidth = setWindowTileWidth \ 2
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
    TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
    
    engineBaseSpeed = engineSpeed
    
    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
    MainViewWidth = TilePixelWidth * WindowTileWidth
    MainViewHeight = TilePixelHeight * WindowTileHeight
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
 
    'Set the view rect
    With MainViewRect
        .Left = MainViewLeft
        .Top = MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    'Set the dest rect
    With MainDestRect
        .Left = TilePixelWidth * TileBufferSize - TilePixelWidth
        .Top = TilePixelHeight * TileBufferSize - TilePixelHeight
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
On Error Resume Next
    Set DirectX = New DirectX7
    
    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If

    
    '****** INIT DirectDraw ******
    ' Create the root DirectDraw object
    Set DirectDraw = DirectX.DirectDrawCreate("")
    
    If Err Then
        MsgBox "No se puede iniciar DirectDraw. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If
    
On Error GoTo 0
    Call DirectDraw.SetCooperativeLevel(setDisplayFormhWnd, DDSCL_NORMAL)
    
    'Primary Surface
    ' Fill the surface description structure
    With SurfaceDesc
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With
    ' Create the surface
    Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)
    
    'Create Primary Clipper
    Set PrimaryClipper = DirectDraw.CreateClipper(0)
    Call PrimaryClipper.SetHWnd(frmMain.hwnd)
    Call PrimarySurface.SetClipper(PrimaryClipper)
    
    With BackBufferRect
        .Left = 0
        .Top = 0
        .Right = TilePixelWidth * (WindowTileWidth + 2 * TileBufferSize)
        .Bottom = TilePixelHeight * (WindowTileHeight + 2 * TileBufferSize)
    End With
    
    With SurfaceDesc
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        If ClientSetup.bUseVideo Then
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
        Else
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        End If
        .lHeight = BackBufferRect.Bottom
        .lWidth = BackBufferRect.Right
    End With
    
    ' Create surface
    Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)
    
    'Set color key
    ddck.low = 0
    ddck.high = 0
    Call BackBufferSurface.SetColorKey(DDCKEY_SRCBLT, ddck)
    
    'Set font transparency
    Call BackBufferSurface.SetFontTransparency(D_TRUE)
    
    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    If MD5File(App.path & "\Recursos\Clases\Clases.ind") <> "b5617f8dc398c34ac499e31b0fffd874" Then
        End
    End If
    Call CargarCascos
    Call CargarFxs
    
    LTLluvia(0) = 224
    LTLluvia(1) = 352
    LTLluvia(2) = 480
    LTLluvia(3) = 608
    LTLluvia(4) = 736
    
    Call LoadGraphics
    
    InitTileEngine = True
End Function

Public Sub DeinitTileEngine()
'***************************************************
'Author: Juan Mart?n Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Destroys all DX objects
'***************************************************
On Error Resume Next
    Set PrimarySurface = Nothing
    Set PrimaryClipper = Nothing
    Set BackBufferSurface = Nothing
    
    Set DirectDraw = Nothing
    
    Set DirectX = Nothing
End Sub

Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)
'***************************************************
'Author: Arron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Mart?n Sotuyo Dodero (Maraxus)
'Updates the game's model and renders everything.
'***************************************************
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    
    '****** Set main view rectangle ******
    MainViewRect.Left = (DisplayFormLeft / Screen.TwipsPerPixelX) + MainViewLeft
    MainViewRect.Top = (DisplayFormTop / Screen.TwipsPerPixelY) + MainViewTop
    MainViewRect.Right = MainViewRect.Left + MainViewWidth
    MainViewRect.Bottom = MainViewRect.Top + MainViewHeight
    
    If EngineRun Then
        If UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                If Not UserMontando Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
            Else
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame * Velocidad / 10 'Reemplacen este 10 por el valor por el que quieran dividir la vel
            End If
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                If Not UserMontando Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
            Else
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame * Velocidad / 10 'Reemplacen este 10 por el valor por el que quieran dividir la vel
            End If
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If
        
        'Update mouse position within view area
        Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
        
        '****** Update screen ******
        If UserCiego Then
            Call CleanViewPort
        Else
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
        End If

        'If IScombate Then Call RenderText(170, 165, "Modo Combate", vbRed, frmMain.font)
        If ClientSetup.bActive Then
        
            If isCapturePending Then
                Call ScreenCapture(True)
                isCapturePending = False
            End If
        End If
        Call Dialogos.Render
        Call DibujarCartel
        
        Call DialogosClanes.Draw
        'Display front-buffer!
        Call PrimarySurface.Blt(MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT)
        
        'Limit FPS to 100 (an easy number higher than monitor's vertical refresh rates)
       Dim TimeSleep As Long ' Declaramos el tiempo para pausear la PC
       If TSetup.bFPS Then
        TimeSleep = (1000 / 60 - 1) - SetElapsedTime(False) ' Cambiar numero 100 para cambiar los FPS
        Else
        TimeSleep = (1000 / 100 - 1) - SetElapsedTime(False)
        End If
        If TimeSleep > 0 Then 'Si el Tiempo es negativo entonces no limitamos
            Sleep TimeSleep
        End If
        Call SetElapsedTime(True) 'Seteamos el tiempo
        
        'FPS update
        If fpsLastCheck + 1000 < DirectX.TickCount Then
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            fpsLastCheck = DirectX.TickCount
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
        
         timerElapsedTime = GetElapsedTime()
            timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
End If
End Sub

Public Function SetElapsedTime(ByVal Start As Boolean) As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 23/05/2011 By MaTeO
'Gets the time that past since the last call
'[MaTeO] Agrego cambios a la funcion
'**************************************************************
    Dim Start_Time As Currency
    Static End_Time As Currency
    Static Timer_Freq As Currency
    'Get the timer frequency
    If Timer_Freq = 0 Then
        QueryPerformanceFrequency Timer_Freq
    End If
   
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
   
    If Not Start Then
        'Calculate elapsed time
        SetElapsedTime = (Start_Time - End_Time) / Timer_Freq * 1000
   
        'Get next end time
    Else
        Call QueryPerformanceCounter(End_Time)
    End If
End Function


Public Sub EfectoNoche(ByRef Surface As DirectDrawSurface7)
    Dim dArray() As Byte
    Dim ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
    Dim rRect As RECT
    
    Surface.GetSurfaceDesc ddsdDest
    
    With rRect
        .Left = 0
        .Top = 0
        .Right = ddsdDest.lWidth
        .Bottom = ddsdDest.lHeight
    End With
    
    If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
        Modo = 0
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
        Modo = 1
    Else
        Modo = 2
    End If
    
    Dim DstLock As Boolean
    DstLock = False
    
    On Local Error GoTo HayErrorAlpha
    
    Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
    DstLock = True
    
    Surface.GetLockedArray dArray()
    Call BltEfectoNoche(ByVal VarPtr(dArray(0, 0)), _
        ddsdDest.lWidth, ddsdDest.lHeight, ddsdDest.lPitch, _
        Modo)
    
HayErrorAlpha:
    If DstLock = True Then
        Surface.Unlock rRect
        DstLock = False
    End If
End Sub

Public Sub RenderText(ByVal lngXPos As Integer, ByVal lngYPos As Integer, ByRef strText As String, ByVal lngColor As Long, ByRef font As StdFont)
    If strText <> "" Then
        Call BackBufferSurface.SetForeColor(vbBlack)
        Call BackBufferSurface.SetFont(font)
        Call BackBufferSurface.DrawText(lngXPos - 2, lngYPos - 1, strText, False)
        
        Call BackBufferSurface.SetForeColor(lngColor)
        Call BackBufferSurface.DrawText(lngXPos, lngYPos, strText, False)
    End If
End Sub

Public Sub RenderTextCentered(ByVal lngXPos As Integer, ByVal lngYPos As Integer, ByRef strText As String, ByVal lngColor As Long, ByRef font As StdFont, Optional ByVal bNegro As Boolean = True)
    Dim hdc As Long
    Dim ret As Size
   
    If strText <> "" Then
        Call BackBufferSurface.SetFont(font)
       
        'Get width of text once rendered
        hdc = BackBufferSurface.GetDC()
        Call GetTextExtentPoint32(hdc, strText, Len(strText), ret)
        Call BackBufferSurface.ReleaseDC(hdc)
       
        lngXPos = lngXPos - ret.cx \ 2
       
    If bNegro Then
            Call BackBufferSurface.SetForeColor(vbBlack)
            Call BackBufferSurface.DrawText(lngXPos - 2, lngYPos - 1, strText, False)
        End If
 
        Call BackBufferSurface.SetForeColor(lngColor)
        Call BackBufferSurface.DrawText(lngXPos, lngYPos, strText, False)
    End If
End Sub

Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim Start_Time As Currency
    Static End_Time As Currency
    Static Timer_Freq As Currency

    'Get the timer frequency
    If Timer_Freq = 0 Then
        QueryPerformanceFrequency Timer_Freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
    
    'Calculate elapsed time
    GetElapsedTime = (Start_Time - End_Time) / Timer_Freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(End_Time)
End Function

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'***************************************************
'Author: Juan Mart?n Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
If Algo > 0 Then
      Algo = Algo - 1
      If ColR > 0 Then
      ColR = ColR - 1
      colG = colG - 1
      colB = colB - 1
      End If
      'frmConnect.font.Name = "Georgia"
      'frmConnect.font.Size = 24
      'RenderTextCentered 500, 300, UserMapName, RGB(ColR / 10, colG / 10, colB / 10), frmConnect.font
      If Algo < 0 Then Algo = 0
      End If
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long
    
    With charlist(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                If Not UserMontando Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
            Else
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame * Velocidad / 10 'Reemplacen este 10 por el valor por el que quieran dividir la vel
            End If
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                If Not UserMontando Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
            Else
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame * Velocidad / 10
            End If
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
       'If done moving stop animation
       If Not moved Then
            If .Heading < 1 Or .Heading > 4 Then .Heading = EAST
            
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
           
            If Not .Movimient Then
           
            .Arma.WeaponWalk(.Heading).Started = 0
            .Arma.WeaponWalk(.Heading).FrameCounter = 1
           
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
           
            End If
           
            .Moving = False
        End If
       
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
       ' If .MinHp > 0 Then
           ' BackBufferSurface.SetForeColor vbBlack
            BackBufferSurface.SetFillColor vbRed
         ' '  Call BackBufferSurface.DrawBox(PixelOffsetX, PixelOffsetY + 50, (((.MinHp / 100) / (.MaxHp / 100)) * 45) + PixelOffsetX, PixelOffsetY + 20)
            
       ' End If
        
        'If .MinMan > 0 Then
            'BackBufferSurface.SetForeColor vbBlack
            'BackBufferSurface.SetFillColor vbCyan
            'Call BackBufferSurface.DrawBox(PixelOffsetX, PixelOffsetY + 50, (((.MinMan / 100) / (.MaxMan / 100)) * 45) + PixelOffsetX, PixelOffsetY + 60)
            
        'End If
  
 
 
            
        If .Head.Head(.Heading).GrhIndex Then
            If Not .invisible Then
            If ConAlfaB = 0 Then
            Movement_Speed = 0.5
                              'Draw Body
                If .Body.Walk(.Heading).GrhIndex Then
                    Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                 End If
                 
                'Draw Head
                If .Head.Head(.Heading).GrhIndex Then
    
                    Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0)
                    
                    'Draw Helmet
                    If .Casco.Head(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0)
                    
                    'Draw Weapon
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                    
                    'Draw Shield
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                
                    'Draw name over head
                    If LenB(.Nombre) > 0 Then
                        If Nombres Then
                            Pos = getTagPosition(.Nombre)
                            'Pos = InStr(.Nombre, "<")
                            'If Pos = 0 Then Pos = Len(.Nombre) + 2
                            
                            If .priv = 0 Then
                                If .Team > 0 Then
                                    Select Case .Team
                                        Case 1
                                            color = RGB(255, 88, 50)
                                        Case 2
                                            color = RGB(60, 190, 37)
                                    End Select
                                    
                                Else
                                    If .Criminal Then
                                        color = RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                    Else
                                        color = RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                    End If
                                End If
                            Else
                                color = RGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            End If
                          
                                                                    
                           
                              
                            'Nick
                           line = Left$(.Nombre, Pos - 2)
                            Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 30, line, color, frmMain.font)
                            
                            color = RGB(255, 130, 0)
                            
                            'Clan
                            line = mid$(.Nombre, Pos)
                            Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 45, line, color, frmMain.font)
                        
                          
End If
           End If
           End If
            End If
            End If
            End If
                         
      If .Head.Head(.Heading).GrhIndex Then
            If Not .invisible Then
            If ConAlfaB = 1 Then
            Movement_Speed = 0.5
            
   If .Body.Walk(.Heading).GrhIndex Then
   If .iBody <> 8 And .iBody <> 145 Then
                    Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                Else
                    DDrawTransGrhtoSurfaceAlpha .Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1
                 End If
                 End If
                 
                'Draw Head
                If .Head.Head(.Heading).GrhIndex Then
                If .iHead <> 500 And .iHead <> 501 Then
                    Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0)
                   Else
                   DDrawTransGrhtoSurfaceAlpha .Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0
                   End If

                    'Draw Helmet
                    If .Casco.Head(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0)
                    
                    'Draw Weapon
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                    
                    'Draw Shield
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                
                    'Draw name over head
                    If LenB(.Nombre) > 0 Then
                        If Nombres Then
                            Pos = getTagPosition(.Nombre)
                            'Pos = InStr(.Nombre, "<")
                            'If Pos = 0 Then Pos = Len(.Nombre) + 2
                            
                            If .priv = 0 Then
                                If .Team > 0 Then
                                    Select Case .Team
                                        Case 1
                                            color = RGB(255, 88, 50)
                                        Case 2
                                            color = RGB(60, 190, 37)
                                    End Select
                                Else
                                    If .Criminal Then
                                        color = RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                    Else
                                        color = RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                    End If
                                End If
                            Else
                                color = RGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            End If
                          
                                                                    
                           
                            'Nick
                           line = Left$(.Nombre, Pos - 2)
                            Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 30, line, color, frmMain.font)
                            
                            'Clan
                            line = mid$(.Nombre, Pos)
                            Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 45, line, color, frmMain.font)
                           
                  End If
           End If
           End If
           End If
                
            ElseIf .invisible Then
            If ConAlfaB = 1 Then
            If esGM(UserCharIndex) = True Then Exit Sub
            If UserOcu = 1 Then Exit Sub
             If .cCont > 0 Then .cCont = .cCont - 1
           If .cCont = 0 Then
            .Drawers = .Drawers + 1
            If .Drawers = 156 Then .Drawers = 0: .cCont = 400
            
                   If .Body.Walk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurfaceAlpha(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
            
                'Draw Head
                If .Head.Head(.Heading).GrhIndex Then
                    Call DDrawTransGrhtoSurfaceAlpha(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0)
                    
                    'Draw Helmet
                    If .Casco.Head(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurfaceAlpha(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0)
                    
                    'Draw Weapon
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurfaceAlpha(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                    
                    'Draw Shield
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurfaceAlpha(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                
                            End If
                            End If
                        

                   
                      ElseIf ConAlfaB = 0 Then
                       If .cCont > 0 Then .cCont = .cCont - 1
                       If esGM(UserCharIndex) = True Then Exit Sub
                       If UserOcu = 1 Then Exit Sub
           If .cCont = 0 Then
            .Drawers = .Drawers + 1
            If .Drawers = 156 Then .Drawers = 0: .cCont = 400
                      If .Body.Walk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
            
                'Draw Head
                If .Head.Head(.Heading).GrhIndex Then
                    Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0)
                    
                    'Draw Helmet
                    If .Casco.Head(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0)
                    
                    'Draw Weapon
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                    
                    'Draw Shield
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
            End If
            End If
            End If
        End If
        
                Else
            'Draw Body
            If .Body.Walk(.Heading).GrhIndex Then _
                Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
        End If
        
        If .Infected = 1 Then
     'Nick infectado
     line = Left$(.Nombre, Pos - 2)
     Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 30, line, RGB(255, 255, 255), frmMain.font)
                           
     'Clan infectado
     line = mid$(.Nombre, Pos)
     Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 45, line, RGB(255, 255, 255), frmMain.font)
       
     'Tagged Infected for user
     line = "      [EVENT BOSS]"
     Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 60, line, RGB(255, 255, 255), frmMain.font)
     End If
        
                If .Angel = 1 Then
     'Nick infectado
     line = Left$(.Nombre, Pos - 2)
     Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 30, line, RGB(255, 255, 255), frmMain.font)
                           
     'Clan infectado
     line = mid$(.Nombre, Pos)
     Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 45, line, RGB(255, 255, 255), frmMain.font)
       
     'Tagged Infected for user
     line = "      [EVENT BOSS]"
     Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 90, line, RGB(255, 255, 255), frmMain.font)
     End If
     
             If .Demonio = 1 Then
     'Nick infectado
     line = Left$(.Nombre, Pos - 2)
     Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 30, line, RGB(255, 255, 255), frmMain.font)
                           
     'Clan infectado
     line = mid$(.Nombre, Pos)
     Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 45, line, RGB(255, 255, 255), frmMain.font)
       
     'Tagged Infected for user
     line = "      [EVENT BOSS]"
     Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 60, line, RGB(255, 255, 255), frmMain.font)
     End If
        
        
          'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, CharIndex)  '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo
        Movement_Speed = 1
        
               'Draw FX
        If .FxIndex <> 0 Then
Dim XDATAFX As Integer, YDATAFX As Integer
        
            '@Nota de Dunkan: Arreglar desde el INDICE.
            If .FxIndex = 1 Then
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
                YDATAFX = PixelOffsetY + 25
            ElseIf .FxIndex = 18 Then
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
                YDATAFX = PixelOffsetY - 15
            ElseIf .FxIndex = 17 Then
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
                YDATAFX = PixelOffsetY - 15
            ElseIf .FxIndex = 19 Then
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
                YDATAFX = PixelOffsetY + 25
            ElseIf .FxIndex = 7 Then    'TORMENTA DE FUEGO
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
                YDATAFX = PixelOffsetY + 30
            ElseIf .FxIndex = 8 Then    'PARALIZAR
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
                YDATAFX = PixelOffsetY + 35
            ElseIf .FxIndex = 9 Then    'CURAR GRAVES
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
                YDATAFX = PixelOffsetY + 25
            ElseIf .FxIndex = 12 Then   'INMO
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
                YDATAFX = PixelOffsetY + 20
            Else
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffsetX)
                YDATAFX = PixelOffsetY
            End If
            If (ConAlfaB = 1) Then
            DDrawTransGrhtoSurfaceAlpha .fX, XDATAFX, YDATAFX, 1, 1
            Else
             Call DDrawTransGrhtoSurface(.fX, XDATAFX, YDATAFX, 1, 1)
           End If
            
            'Check if animation is over
            If .fX.Started = 0 Then _
                .FxIndex = 0
        End If
    End With
End Sub


Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Mart?n Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With charlist(CharIndex)
        .FxIndex = fX
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
        
            .fX.Loops = Loops
        End If
    End With
End Sub

Private Sub CleanViewPort()
'***************************************************
'Author: Juan Mart?n Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Fills the viewport with black.
'***************************************************
    Dim r As RECT
    Call BackBufferSurface.BltColorFill(r, vbBlack)
End Sub
Public Function CharTieneClan() As Boolean

Dim tPos As Integer

tPos = InStr(charlist(UserCharIndex).Nombre, "<")

If tPos = 0 Then
CharTieneClan = False
Exit Function
End If

CharTieneClan = True

End Function
Public Function CharTieneParty() As Boolean

Dim tPos As Integer

tPos = InStr(charlist(UserCharIndex).Nombre, "<")

If tPos = 0 Then
CharTieneParty = False
Exit Function
End If

CharTieneParty = True

End Function

Sub DDrawTransGrhtoSurfaceAlphaTecho(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, ByVal Alpha As Byte)
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]
 
 
    If Animate Then
        If Grh.Started = 1 Then
        Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * Movement_Speed
         
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = 1
             
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
 
'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
Dim QuitarAnimacion As Boolean
 
If Grh.GrhIndex = 0 Then Exit Sub
 
 
'Figure out what frame to draw (always 1 if not animated)
If Grh.FrameCounter < 1 Then Grh.FrameCounter = 1
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
 
'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If
 
With SourceRect
    .Left = GrhData(iGrhIndex).sX + IIf(X < 0, Abs(X), 0)
    .Top = GrhData(iGrhIndex).sY + IIf(Y < 0, Abs(Y), 0)
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With
 
'surface.BltFast X, Y, SurfaceDB.surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
 
Dim Src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim Modo As Long
 
Set Src = SurfaceDB.Surface(GrhData(iGrhIndex).FileNum)
 
Src.GetSurfaceDesc ddsdSrc
Surface.GetSurfaceDesc ddsdDest
 
With rDest
    .Left = X
    .Top = Y
    .Right = X + GrhData(iGrhIndex).pixelWidth
    .Bottom = Y + GrhData(iGrhIndex).pixelHeight
 
    If .Right > ddsdDest.lWidth Then
        .Right = ddsdDest.lWidth
    End If
    If .Bottom > ddsdDest.lHeight Then
        .Bottom = ddsdDest.lHeight
    End If
End With
 
' 0 -> 16 bits 555
' 1 -> 16 bits 565
' 2 -> 16 bits raro (Sin implementar)
' 3 -> 24 bits
' 4 -> 32 bits
 
If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 555
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 565
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 565
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = 65280 And ddsdSrc.ddpfPixelFormat.lGBitMask = 65280 Then
    Modo = 565
Else
    'Modo = 2 '16 bits raro ?
    Surface.BltFast X, Y, Src, SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Exit Sub
End If
 
Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False
 
On Local Error GoTo HayErrorAlpha
 
Src.Lock SourceRect, ddsdSrc, DDLOCK_WAIT, 0
SrcLock = True
Surface.Lock rDest, ddsdDest, DDLOCK_WAIT, 0
DstLock = True
 
Surface.GetLockedArray dArray()
Src.GetLockedArray sArray()
 
Call vbDABLalphablend16(Modo, 1, ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), ByVal VarPtr(dArray(X + X, Y)), Alpha, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 0)
 
Surface.Unlock rDest
DstLock = False
Src.Unlock SourceRect
SrcLock = False
 
 
Exit Sub
 
HayErrorAlpha:
If SrcLock Then Src.Unlock SourceRect
If DstLock Then Surface.Unlock rDest
 
End Sub
 
