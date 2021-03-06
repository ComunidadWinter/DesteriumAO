Attribute VB_Name = "Extra"
'Argentum Online 0.12.2
'Copyright (C) 2002 M?rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n?mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C?digo Postal 1900
'Pablo Ignacio M?rquez

Option Explicit

Public Function EsNewbie(ByVal Userindex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    EsNewbie = UserList(Userindex).Stats.ELV <= LimiteNewbie
End Function
Public Function esArmada(ByVal Userindex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************

    esArmada = (UserList(Userindex).Faccion.ArmadaReal = 1)
End Function

Public Function esCaos(ByVal Userindex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************

    esCaos = (UserList(Userindex).Faccion.FuerzasCaos = 1)
End Function

Public Function EsGM(ByVal Userindex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************

    EsGM = (UserList(Userindex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero))
End Function
Public Function EsVip(ByVal Userindex As Integer) As Boolean 'es un usuario vip?
EsVip = UserList(Userindex).flags.Oro
End Function
Public Function EsVipp(ByVal Userindex As Integer) As Boolean 'es un usuario vip?
EsVipp = UserList(Userindex).flags.Plata
End Function
Public Function EsVipb(ByVal Userindex As Integer) As Boolean 'es un usuario vip?
EsVipb = UserList(Userindex).flags.Bronce
End Function
Public Function EsCuarenta(ByVal Userindex As Integer) As Boolean
    EsCuarenta = UserList(Userindex).Stats.ELV >= 40
End Function
Public Function EsSiete(ByVal Userindex As Integer) As Boolean
    EsSiete = UserList(Userindex).Stats.ELV >= 47
End Function
Public Function EsOcho(ByVal Userindex As Integer) As Boolean
    EsOcho = UserList(Userindex).Stats.ELV >= 48
End Function
Public Function EsNueve(ByVal Userindex As Integer) As Boolean
    EsNueve = UserList(Userindex).Stats.ELV >= 49
End Function
Public Function EsQuince(ByVal Userindex As Integer) As Boolean
    EsQuince = UserList(Userindex).Stats.ELV >= 15
End Function
Public Function EsVeinte(ByVal Userindex As Integer) As Boolean
    EsVeinte = UserList(Userindex).Stats.ELV >= 20
End Function
Public Function EsVeinticinco(ByVal Userindex As Integer) As Boolean
    EsVeinticinco = UserList(Userindex).Stats.ELV >= 25
End Function
Public Function EsQuinceM(ByVal Userindex As Integer) As Boolean
    EsQuinceM = UserList(Userindex).Stats.ELV <= 15
End Function
Public Function EsTreintaM(ByVal Userindex As Integer) As Boolean
    EsTreintaM = UserList(Userindex).Stats.ELV >= 13
End Function
Public Function EsHM(ByVal Userindex As Integer) As Boolean
    EsHM = UserList(Userindex).Stats.ELV >= 30
End Function
Public Function EsUM(ByVal Userindex As Integer) As Boolean
    EsUM = UserList(Userindex).Stats.ELV >= 35
End Function
Public Function EsMM(ByVal Userindex As Integer) As Boolean
    EsMM = UserList(Userindex).Stats.ELV >= 45
End Function
Public Function NoEsUM(ByVal Userindex As Integer) As Boolean 'es un usuario vip?
NoEsUM = UserList(Userindex).flags.Oro <= 0
End Function

Public Sub DoTileEvents(ByVal Userindex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 06/03/2010
'Handles the Map passage of Users. Allows the existance
'of exclusive maps for Newbies, Royal Army and Caos Legion members
'and enables GMs to enter every map without restriction.
'Uses: Mapinfo(map).Restringir = "NEWBIE" (newbies), "ARMADA", "CAOS", "FACCION" or "NO".
' 06/03/2010 : Now we have 5 attemps to not fall into a map change or another teleport while going into a teleport. (Marco)
'***************************************************

    Dim nPos As WorldPos
    Dim FxFlag As Boolean
    Dim TelepRadio As Integer
    Dim DestPos As WorldPos
    
On Error GoTo Errhandler
    'Controla las salidas
    If InMapBounds(map, X, Y) Then
        With MapData(map, X, Y)
            If .ObjInfo.objindex > 0 Then
                FxFlag = ObjData(.ObjInfo.objindex).OBJType = eOBJType.otTeleport
                TelepRadio = ObjData(.ObjInfo.objindex).Radio
            End If
            
            If .TileExit.map > 0 And .TileExit.map <= NumMaps Then
                
                ' Es un teleport, entra en una posicion random, acorde al radio (si es 0, es pos fija)
                ' We have 5 attempts to not falling into another teleport or a map exit.. If we get to the fifth attemp,
                ' the teleport will act as if its radius = 0.
                If FxFlag And TelepRadio > 0 Then
                    Dim attemps As Long
                    Dim exitMap As Boolean
                    Do
                        DestPos.X = .TileExit.X + RandomNumber(TelepRadio * (-1), TelepRadio)
                        DestPos.Y = .TileExit.Y + RandomNumber(TelepRadio * (-1), TelepRadio)
                        
                        attemps = attemps + 1
                        
                        exitMap = MapData(.TileExit.map, DestPos.X, DestPos.Y).TileExit.map > 0 And _
                                MapData(.TileExit.map, DestPos.X, DestPos.Y).TileExit.map <= NumMaps
                    Loop Until (attemps >= 5 Or exitMap = False)
                    
                    If attemps >= 5 Then
                        DestPos.X = .TileExit.X
                        DestPos.Y = .TileExit.Y
                    End If
                ' Posicion fija
                Else
                    DestPos.X = .TileExit.X
                    DestPos.Y = .TileExit.Y
                End If
                
                DestPos.map = .TileExit.map
                
                If EsGM(Userindex) Then
                    Call LogGM(UserList(Userindex).Name, "Utiliz? un teleport hacia el mapa " & _
                        DestPos.map & " (" & DestPos.X & "," & DestPos.Y & ")")
                End If
                
                '?Es mapa de newbies?
                If UCase$(MapInfo(DestPos.map).Restringir) = "NEWBIE" Then
                    '?El usuario es un newbie?
                    If EsNewbie(Userindex) Or EsGM(Userindex) Then
                        If LegalPos(DestPos.map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
                            Call WarpUserChar(Userindex, DestPos.map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es newbie
                        Call WriteConsoleMsg(Userindex, "Mapa exclusivo para newbies.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(Userindex).Pos, nPos)
        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
                        End If
                    End If
                ElseIf UCase$(MapInfo(DestPos.map).Restringir) = "ARMADA" Then '?Es mapa de Armadas?
                    '?El usuario es Armada?
                    If esArmada(Userindex) Or EsGM(Userindex) Then
                        If LegalPos(DestPos.map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
                            Call WarpUserChar(Userindex, DestPos.map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es armada
                        Call WriteConsoleMsg(Userindex, "Mapa exclusivo para miembros del ej?rcito real.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(Userindex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                ElseIf UCase$(MapInfo(DestPos.map).Restringir) = "CAOS" Then '?Es mapa de Caos?
                    '?El usuario es Caos?
                    If esCaos(Userindex) Or EsGM(Userindex) Then
                        If LegalPos(DestPos.map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
                            Call WarpUserChar(Userindex, DestPos.map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es caos
                        Call WriteConsoleMsg(Userindex, "Mapa exclusivo para miembros de la legi?n oscura.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(Userindex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                ElseIf UCase$(MapInfo(DestPos.map).Restringir) = "FACCION" Then '?Es mapa de faccionarios?
                    '?El usuario es Armada o Caos?
                    If esArmada(Userindex) Or esCaos(Userindex) Or EsGM(Userindex) Then
                        If LegalPos(DestPos.map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
                            Call WarpUserChar(Userindex, DestPos.map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es Faccionario
                        Call WriteConsoleMsg(Userindex, "Solo se permite entrar al mapa si eres miembro de alguna facci?n.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(Userindex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
 
                    'QUince
                   
    ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "CUARENTA" Then
                           If UserList(Userindex).Stats.ELV >= 40 Then
                            If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
                                Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
                            Else
                                Call ClosestLegalPos(.TileExit, nPos)
                               If nPos.X <> 0 And nPos.Y <> 0 Then
                                    Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                                End If
                           End If
                       Else
                           Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
                           Call ClosestStablePos(UserList(Userindex).Pos, nPos)
 
                           If nPos.X <> 0 And nPos.Y <> 0 Then
                               Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
                           End If
                       End If
                       
                       
               ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "MENOSCINCO" Then
                           If UserList(Userindex).Stats.ELV <= 45 Then
                            If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
                                Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
                            Else
                                Call ClosestLegalPos(.TileExit, nPos)
                               If nPos.X <> 0 And nPos.Y <> 0 Then
                                    Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                                End If
                           End If
                       Else
                           Call WriteConsoleMsg(Userindex, "Tu nivel es muy elevado para ingresar en este mapa.", FontTypeNames.FONTTYPE_INFO)
                           Call ClosestStablePos(UserList(Userindex).Pos, nPos)
 
                           If nPos.X <> 0 And nPos.Y <> 0 Then
                               Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
                           End If
                       End If
                       
                ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "MENOSCUATRO" Then
                           If UserList(Userindex).Stats.ELV <= 40 Then
                            If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
                                Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
                            Else
                                Call ClosestLegalPos(.TileExit, nPos)
                               If nPos.X <> 0 And nPos.Y <> 0 Then
                                    Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                                End If
                           End If
                       Else
                           Call WriteConsoleMsg(Userindex, "Tu nivel es muy elevado para ingresar en este mapa.", FontTypeNames.FONTTYPE_INFO)
                           Call ClosestStablePos(UserList(Userindex).Pos, nPos)
 
                           If nPos.X <> 0 And nPos.Y <> 0 Then
                               Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
                           End If
                       End If
                       
                        ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "CINCO" Then
                           If UserList(Userindex).Stats.ELV >= 45 Then
                            If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
                                Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
                            Else
                                Call ClosestLegalPos(.TileExit, nPos)
                               If nPos.X <> 0 And nPos.Y <> 0 Then
                                    Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                                End If
                           End If
                       Else
                           Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
                           Call ClosestStablePos(UserList(Userindex).Pos, nPos)
 
                           If nPos.X <> 0 And nPos.Y <> 0 Then
                               Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
                           End If
                       End If
                       
                        ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "SEIS" Then
                           If UserList(Userindex).Stats.ELV >= 46 Then
                            If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
                                Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
                            Else
                                Call ClosestLegalPos(.TileExit, nPos)
                               If nPos.X <> 0 And nPos.Y <> 0 Then
                                    Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                                End If
                           End If
                       Else
                           Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
                           Call ClosestStablePos(UserList(Userindex).Pos, nPos)
 
                           If nPos.X <> 0 And nPos.Y <> 0 Then
                               Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
                           End If
                       End If
                    
        ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "SIETE" Then
                           If UserList(Userindex).Stats.ELV >= 47 Then
                            If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
                                Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
                            Else
                                Call ClosestLegalPos(.TileExit, nPos)
                               If nPos.X <> 0 And nPos.Y <> 0 Then
                                    Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                                End If
                           End If
                       Else
                           Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
                           Call ClosestStablePos(UserList(Userindex).Pos, nPos)
 
                           If nPos.X <> 0 And nPos.Y <> 0 Then
                               Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
                           End If
                       End If
                       
            ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "OCHO" Then
                           If UserList(Userindex).Stats.ELV >= 48 Then
                            If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
                                Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
                            Else
                                Call ClosestLegalPos(.TileExit, nPos)
                               If nPos.X <> 0 And nPos.Y <> 0 Then
                                    Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                                End If
                           End If
                       Else
                           Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
                           Call ClosestStablePos(UserList(Userindex).Pos, nPos)
 
                           If nPos.X <> 0 And nPos.Y <> 0 Then
                               Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
                           End If
                       End If
                       
                       
                ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "NUEVE" Then
                           If UserList(Userindex).Stats.ELV >= 49 Then
                            If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
                                Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
                            Else
                                Call ClosestLegalPos(.TileExit, nPos)
                               If nPos.X <> 0 And nPos.Y <> 0 Then
                                    Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                                End If
                           End If
                       Else
                           Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
                           Call ClosestStablePos(UserList(Userindex).Pos, nPos)
 
                           If nPos.X <> 0 And nPos.Y <> 0 Then
                               Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
                           End If
                       End If
 
  'QUInce FIN
  
  
      ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "QUINCE" Then
                           If UserList(Userindex).Stats.ELV >= 15 Then
                            If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
                                Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
                            Else
                                Call ClosestLegalPos(.TileExit, nPos)
                               If nPos.X <> 0 And nPos.Y <> 0 Then
                                    Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                                End If
                           End If
                       Else
                           Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
                           Call ClosestStablePos(UserList(Userindex).Pos, nPos)
 
                           If nPos.X <> 0 And nPos.Y <> 0 Then
                               Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
                           End If
                       End If
  
  
        ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "VEINTE" Then
                           If UserList(Userindex).Stats.ELV >= 20 Then
                            If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
                                Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
                            Else
                                Call ClosestLegalPos(.TileExit, nPos)
                               If nPos.X <> 0 And nPos.Y <> 0 Then
                                    Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                                End If
                           End If
                       Else
                           Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
                           Call ClosestStablePos(UserList(Userindex).Pos, nPos)
 
                           If nPos.X <> 0 And nPos.Y <> 0 Then
                               Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
                           End If
                       End If
                       
                               ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "VEINTICINCO" Then
                           If UserList(Userindex).Stats.ELV >= 25 Then
                            If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
                                Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
                            Else
                                Call ClosestLegalPos(.TileExit, nPos)
                               If nPos.X <> 0 And nPos.Y <> 0 Then
                                    Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                                End If
                           End If
                       Else
                           Call WriteConsoleMsg(Userindex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
                           Call ClosestStablePos(UserList(Userindex).Pos, nPos)
 
                           If nPos.X <> 0 And nPos.Y <> 0 Then
                               Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
                           End If
                       End If
  
 
 
 ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "VIP" Then
If UserList(Userindex).flags.Oro = 1 Then
If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
Else
Call ClosestLegalPos(.TileExit, nPos)
If nPos.X <> 0 And nPos.Y <> 0 Then
Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
End If
End If
Else
Call WriteConsoleMsg(Userindex, "Solo los usuarios Oro pueden ingresar a este mapa.", FontTypeNames.FONTTYPE_CONSE)
Call ClosestStablePos(UserList(Userindex).Pos, nPos)

If nPos.X <> 0 And nPos.Y <> 0 Then
Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
End If
End If

 ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "PREMIUM" Then
If UserList(Userindex).flags.Premium = 1 Then
If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
Else
Call ClosestLegalPos(.TileExit, nPos)
If nPos.X <> 0 And nPos.Y <> 0 Then
Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
End If
End If
Else
Call WriteConsoleMsg(Userindex, "S?lo los usuarios PREMIUM pueden ingresar a este mapa.", FontTypeNames.FONTTYPE_CONSE)
Call ClosestStablePos(UserList(Userindex).Pos, nPos)

If nPos.X <> 0 And nPos.Y <> 0 Then
Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
End If
End If

 ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "VIPP" Then
If UserList(Userindex).flags.Plata = 1 Or UserList(Userindex).flags.Bronce = 1 Or UserList(Userindex).flags.Oro = 1 Then
If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
Else
Call ClosestLegalPos(.TileExit, nPos)
If nPos.X <> 0 And nPos.Y <> 0 Then
Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
End If
End If
Else
Call WriteConsoleMsg(Userindex, "S?lo los usuarios Bronce, Plata y Oro pueden ingresar a este mapa.", FontTypeNames.FONTTYPE_CONSE)
Call ClosestStablePos(UserList(Userindex).Pos, nPos)

If nPos.X <> 0 And nPos.Y <> 0 Then
Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
End If
End If

 ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "NOESUM" Then
If UserList(Userindex).flags.Oro <= 0 Then
If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
Else
Call ClosestLegalPos(.TileExit, nPos)
If nPos.X <> 0 And nPos.Y <> 0 Then
Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
End If
End If
Else
Call WriteConsoleMsg(Userindex, "Los usuarios Oro no pueden ingresar a este mapa.", FontTypeNames.FONTTYPE_CONSE)
Call ClosestStablePos(UserList(Userindex).Pos, nPos)

If nPos.X <> 0 And nPos.Y <> 0 Then
Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
End If
End If

 
  'QUInce FIN
                    ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "PREMIUM" Then
If UserList(Userindex).flags.Premium = 1 Then
If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(Userindex)) Then
Call WarpUserChar(Userindex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
Else
Call ClosestLegalPos(.TileExit, nPos)
If nPos.X <> 0 And nPos.Y <> 0 Then
Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
End If
End If
Else
Call WriteConsoleMsg(Userindex, "Para ingresar a este mapa tienes que ser un Usuario PREMIUM.", FontTypeNames.FONTTYPE_INFO)
Call ClosestStablePos(UserList(Userindex).Pos, nPos)

If nPos.X <> 0 And nPos.Y <> 0 Then
Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, False)
End If
End If
                Else 'No es un mapa de newbies, ni Armadas, ni Caos, ni faccionario.
                    If LegalPos(DestPos.map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
                        Call WarpUserChar(Userindex, DestPos.map, DestPos.X, DestPos.Y, FxFlag)
                    Else
                        Call ClosestLegalPos(DestPos, nPos)
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(Userindex, nPos.map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                End If
                
                'Te fusite del mapa. La criatura ya no es m?s tuya ni te reconoce como que vos la atacaste.
                Dim aN As Integer
                
                aN = UserList(Userindex).flags.AtacadoPorNpc
                If aN > 0 Then
                   Npclist(aN).Movement = Npclist(aN).flags.OldMovement
                   Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
                   Npclist(aN).flags.AttackedBy = vbNullString
                End If
            
                aN = UserList(Userindex).flags.NPCAtacado
                If aN > 0 Then
                    If Npclist(aN).flags.AttackedFirstBy = UserList(Userindex).Name Then
                        Npclist(aN).flags.AttackedFirstBy = vbNullString
                    End If
                End If
                UserList(Userindex).flags.AtacadoPorNpc = 0
                UserList(Userindex).flags.NPCAtacado = 0
            End If
        End With
    End If
Exit Sub

Errhandler:
    Call LogError("Error en DotileEvents. Error: " & Err.Number & " - Desc: " & Err.Description)
End Sub

Function InRangoVision(ByVal Userindex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If X > UserList(Userindex).Pos.X - MinXBorder And X < UserList(Userindex).Pos.X + MinXBorder Then
        If Y > UserList(Userindex).Pos.Y - MinYBorder And Y < UserList(Userindex).Pos.Y + MinYBorder Then
            InRangoVision = True
            Exit Function
        End If
    End If
    InRangoVision = False

End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, X As Integer, Y As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If X > Npclist(NpcIndex).Pos.X - MinXBorder And X < Npclist(NpcIndex).Pos.X + MinXBorder Then
        If Y > Npclist(NpcIndex).Pos.Y - MinYBorder And Y < Npclist(NpcIndex).Pos.Y + MinYBorder Then
            InRangoVisionNPC = True
            Exit Function
        End If
    End If
    InRangoVisionNPC = False

End Function


Function InMapBounds(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If (map <= 0 Or map > NumMaps) Or X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        InMapBounds = False
    Else
        InMapBounds = True
    End If
    
    End Function

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos, Optional PuedeAgua As Boolean = False, Optional PuedeTierra As Boolean = True)
'*****************************************************************
'Author: Unknown (original version)
'Last Modification: 24/01/2007 (ToxicWaste)
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Long
Dim tY As Long

nPos.map = Pos.map

Do While Not LegalPos(Pos.map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
            
            If LegalPos(nPos.map, tX, tY, PuedeAgua, PuedeTierra) Then
                nPos.X = tX
                nPos.Y = tY
                '?Hay objeto?
                
                tX = Pos.X + LoopC
                tY = Pos.Y + LoopC
            End If
        Next tX
    Next tY
    
    LoopC = LoopC + 1
Loop

If Notfound = True Then
    nPos.X = 0
    nPos.Y = 0
End If

End Sub

Public Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
'***************************************************
'Author: Unknown
'Last Modification: -
'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
'*****************************************************************

    Dim Notfound As Boolean
    Dim LoopC As Integer
    Dim tX As Long
    Dim tY As Long
    
    nPos.map = Pos.map
    
    Do While Not LegalPos(Pos.map, nPos.X, nPos.Y)
        If LoopC > 12 Then
            Notfound = True
            Exit Do
        End If
        
        For tY = Pos.Y - LoopC To Pos.Y + LoopC
            For tX = Pos.X - LoopC To Pos.X + LoopC
                
                If LegalPos(nPos.map, tX, tY) And MapData(nPos.map, tX, tY).TileExit.map = 0 Then
                    nPos.X = tX
                    nPos.Y = tY
                    '?Hay objeto?
                    
                    tX = Pos.X + LoopC
                    tY = Pos.Y + LoopC
                End If
            Next tX
        Next tY
        
        LoopC = LoopC + 1
    Loop
    
    If Notfound = True Then
        nPos.X = 0
        nPos.Y = 0
    End If

End Sub

Function NameIndex(ByVal Name As String) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim Userindex As Long
    
    '?Nombre valido?
    If LenB(Name) = 0 Then
        NameIndex = 0
        Exit Function
    End If
    
    If InStrB(Name, "+") <> 0 Then
        Name = UCase$(Replace(Name, "+", " "))
    End If
    
    Userindex = 1
    Do Until UCase$(UserList(Userindex).Name) = UCase$(Name)
        
        Userindex = Userindex + 1
        
        If Userindex > MaxUsers Then
            NameIndex = 0
            Exit Function
        End If
    Loop
     
    NameIndex = Userindex
End Function
Function CheckForSameHD(ByVal Userindex As Integer, ByVal UserHD As String) As Boolean '//Disco.
 
    Dim LoopC As Long
   
    For LoopC = 1 To MaxUsers
        If UserList(LoopC).flags.UserLogged = True Then
            If UserList(LoopC).HD = UserHD And Userindex <> LoopC Then
                CheckForSameHD = True
                Exit Function
            End If
        End If
    Next LoopC
   
    CheckForSameHD = False
End Function

Function CheckForSameIP(ByVal Userindex As Integer, ByVal UserIP As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim LoopC As Long
    
    For LoopC = 1 To MaxUsers
        If UserList(LoopC).flags.UserLogged = True Then
            If UserList(LoopC).ip = UserIP And Userindex <> LoopC Then
                CheckForSameIP = True
                Exit Function
            End If
        End If
    Next LoopC
    
    CheckForSameIP = False
End Function

Function CheckForSameName(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'Controlo que no existan usuarios con el mismo nombre
    Dim LoopC As Long
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged Then
            
            'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
            'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
            'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
            'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
            'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
            
            If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
                CheckForSameName = True
                Exit Function
            End If
        End If
    Next LoopC
    
    CheckForSameName = False
End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)
'***************************************************
'Author: Unknown
'Last Modification: -
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************

    Select Case Head
        Case eHeading.NORTH
            Pos.Y = Pos.Y - 1
        
        Case eHeading.SOUTH
            Pos.Y = Pos.Y + 1
        
        Case eHeading.EAST
            Pos.X = Pos.X + 1
        
        Case eHeading.WEST
            Pos.X = Pos.X - 1
    End Select
End Sub

Function LegalPos(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal CheckExitTile As Boolean = False) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Checks if the position is Legal.
'***************************************************

    '?Es un mapa valido?
    If (map <= 0 Or map > NumMaps) Or _
       (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
                LegalPos = False
    Else
        With MapData(map, X, Y)
            If PuedeAgua And PuedeTierra Then
                LegalPos = (.Blocked <> 1) And _
                           (.Userindex = 0) And _
                           (.NpcIndex = 0)
            ElseIf PuedeTierra And Not PuedeAgua Then
                LegalPos = (.Blocked <> 1) And _
                           (.Userindex = 0) And _
                           (.NpcIndex = 0) And _
                           (Not HayAgua(map, X, Y))
            ElseIf PuedeAgua And Not PuedeTierra Then
                LegalPos = (.Blocked <> 1) And _
                           (.Userindex = 0) And _
                           (.NpcIndex = 0) And _
                           (HayAgua(map, X, Y))
            Else
                LegalPos = False
            End If
        End With
        
        If CheckExitTile Then
            LegalPos = LegalPos And (MapData(map, X, Y).TileExit.map = 0)
        End If
        
    End If

End Function

Function MoveToLegalPos(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True) As Boolean
'***************************************************
'Autor: ZaMa
'Last Modification: 13/07/2009
'Checks if the position is Legal, but considers that if there's a casper, it's a legal movement.
'13/07/2009: ZaMa - Now it's also legal move where an invisible admin is.
'***************************************************

Dim Userindex As Integer
Dim IsDeadChar As Boolean
Dim IsAdminInvisible As Boolean


'?Es un mapa valido?
If (map <= 0 Or map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
            MoveToLegalPos = False
    Else
        Userindex = MapData(map, X, Y).Userindex
        If Userindex > 0 Then
            IsDeadChar = UserList(Userindex).flags.Muerto = 1
            IsAdminInvisible = UserList(Userindex).flags.AdminInvisible = 1
        Else
            IsDeadChar = False
            IsAdminInvisible = False
        End If
    
    If PuedeAgua And PuedeTierra Then
        MoveToLegalPos = (MapData(map, X, Y).Blocked <> 1) And _
                   (Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                   (MapData(map, X, Y).NpcIndex = 0)
    ElseIf PuedeTierra And Not PuedeAgua Then
        MoveToLegalPos = (MapData(map, X, Y).Blocked <> 1) And _
                   (Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                   (MapData(map, X, Y).NpcIndex = 0) And _
                   (Not HayAgua(map, X, Y))
    ElseIf PuedeAgua And Not PuedeTierra Then
        MoveToLegalPos = (MapData(map, X, Y).Blocked <> 1) And _
                   (Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                   (MapData(map, X, Y).NpcIndex = 0) And _
                   (HayAgua(map, X, Y))
    Else
        MoveToLegalPos = False
    End If
  
End If

End Function
Public Sub FindLegalPos(ByVal Userindex As Integer, ByVal map As Integer, ByRef X As Integer, ByRef Y As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 26/03/2009
'Search for a Legal pos for the user who is being teleported.
'***************************************************

    If MapData(map, X, Y).Userindex <> 0 Or _
        MapData(map, X, Y).NpcIndex <> 0 Then
                    
        ' Se teletransporta a la misma pos a la que estaba
        If MapData(map, X, Y).Userindex = Userindex Then Exit Sub
                            
        Dim FoundPlace As Boolean
        Dim tX As Long
        Dim tY As Long
        Dim Rango As Long
        Dim OtherUserIndex As Integer
    
        For Rango = 1 To 5
            For tY = Y - Rango To Y + Rango
                For tX = X - Rango To X + Rango
                    'Reviso que no haya User ni NPC
                    If MapData(map, tX, tY).Userindex = 0 And _
                        MapData(map, tX, tY).NpcIndex = 0 Then
                        
                        If InMapBounds(map, tX, tY) Then FoundPlace = True
                        
                        Exit For
                    End If

                Next tX
        
                If FoundPlace Then _
                    Exit For
            Next tY
            
            If FoundPlace Then _
                    Exit For
        Next Rango

    
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            X = tX
            Y = tY
        Else
            'Muy poco probable, pero..
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            OtherUserIndex = MapData(map, X, Y).Userindex
            If OtherUserIndex <> 0 Then
                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If UserList(OtherUserIndex).ComUsu.DestUsu > 0 Then
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(OtherUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        Call FinComerciarUsu(UserList(OtherUserIndex).ComUsu.DestUsu)
                        Call WriteConsoleMsg(UserList(OtherUserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                        Call FlushBuffer(UserList(OtherUserIndex).ComUsu.DestUsu)
                    End If
                    'Lo sacamos.
                    If UserList(OtherUserIndex).flags.UserLogged Then
                        Call FinComerciarUsu(OtherUserIndex)
                        Call WriteErrorMsg(OtherUserIndex, "Alguien se ha conectado donde te encontrabas, por favor recon?ctate...")
                        Call FlushBuffer(OtherUserIndex)
                    End If
                End If
            
                Call CloseSocket(OtherUserIndex)
            End If
        End If
    End If

End Sub

Function LegalPosNPC(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte, Optional ByVal IsPet As Boolean = False) As Boolean
'***************************************************
'Autor: Unkwnown
'Last Modification: 09/23/2009
'Checks if it's a Legal pos for the npc to move to.
'09/23/2009: Pato - If UserIndex is a AdminInvisible, then is a legal pos.
'***************************************************
Dim IsDeadChar As Boolean
Dim Userindex As Integer
Dim IsAdminInvisible As Boolean
    
    
    If (map <= 0 Or map > NumMaps) Or _
        (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPosNPC = False
        Exit Function
    End If

    With MapData(map, X, Y)
        Userindex = .Userindex
        If Userindex > 0 Then
            IsDeadChar = UserList(Userindex).flags.Muerto = 1
            IsAdminInvisible = (UserList(Userindex).flags.AdminInvisible = 1)
        Else
            IsDeadChar = False
            IsAdminInvisible = False
        End If
    
        If AguaValida = 0 Then
            LegalPosNPC = (.Blocked <> 1) And _
            (.Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And _
            (.NpcIndex = 0) And _
            (.trigger <> eTrigger.POSINVALIDA Or IsPet) _
            And Not HayAgua(map, X, Y)
        Else
            LegalPosNPC = (.Blocked <> 1) And _
            (.Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And _
            (.NpcIndex = 0) And _
            (.trigger <> eTrigger.POSINVALIDA Or IsPet)
        End If
    End With
End Function


Sub SendHelp(ByVal index As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim NumHelpLines As Integer
Dim LoopC As Integer

NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

For LoopC = 1 To NumHelpLines
    Call WriteConsoleMsg(index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), FontTypeNames.FONTTYPE_INFO)
Next LoopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If Npclist(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(Npclist(NpcIndex).Expresiones(randomi), Npclist(NpcIndex).Char.CharIndex, vbWhite))
    End If
End Sub

Sub LookatTile(ByVal Userindex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 26/03/2009
'13/02/2009: ZaMa - EL nombre del gm que aparece por consola al clickearlo, tiene el color correspondiente a su rango
'***************************************************

On Error GoTo Errhandler

'Responde al click del usuario sobre el mapa
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String
Dim Ft As FontTypeNames
Dim UserName As String

With UserList(Userindex)
    '?Rango Visi?n? (ToxicWaste)
    If (Abs(.Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(.Pos.X - X) > RANGO_VISION_X) Then
        Exit Sub
    End If
    
    '?Posicion valida?
    If InMapBounds(map, X, Y) Then
        With .flags
            .TargetMap = map
            .TargetX = X
            .TargetY = Y
            '?Es un obj?
            If MapData(map, X, Y).ObjInfo.objindex > 0 Then
                'Informa el nombre
                .TargetObjMap = map
                .TargetObjX = X
                .TargetObjY = Y
                FoundSomething = 1
            ElseIf MapData(map, X + 1, Y).ObjInfo.objindex > 0 Then
                'Informa el nombre
                If ObjData(MapData(map, X + 1, Y).ObjInfo.objindex).OBJType = eOBJType.otPuertas Then
                    .TargetObjMap = map
                    .TargetObjX = X + 1
                    .TargetObjY = Y
                    FoundSomething = 1
                End If
            ElseIf MapData(map, X + 1, Y + 1).ObjInfo.objindex > 0 Then
                If ObjData(MapData(map, X + 1, Y + 1).ObjInfo.objindex).OBJType = eOBJType.otPuertas Then
                    'Informa el nombre
                    .TargetObjMap = map
                    .TargetObjX = X + 1
                    .TargetObjY = Y + 1
                    FoundSomething = 1
                End If
            ElseIf MapData(map, X, Y + 1).ObjInfo.objindex > 0 Then
                If ObjData(MapData(map, X, Y + 1).ObjInfo.objindex).OBJType = eOBJType.otPuertas Then
                    'Informa el nombre
                    .TargetObjMap = map
                    .TargetObjX = X
                    .TargetObjY = Y + 1
                    FoundSomething = 1
                End If
            End If
            
            If FoundSomething = 1 Then
                .TargetObj = MapData(map, .TargetObjX, .TargetObjY).ObjInfo.objindex
                If MostrarCantidad(.TargetObj) Then
                    Call WriteConsoleMsg(Userindex, ObjData(.TargetObj).Name & " - " & MapData(.TargetObjMap, .TargetObjX, .TargetObjY).ObjInfo.Amount & "", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, ObjData(.TargetObj).Name, FontTypeNames.FONTTYPE_INFO)
                End If
            
            End If
             '?Es un personaje?
            If Y + 1 <= YMaxMapSize Then
                If MapData(map, X, Y + 1).Userindex > 0 Then
                    TempCharIndex = MapData(map, X, Y + 1).Userindex
                    FoundChar = 1
                End If
                If MapData(map, X, Y + 1).NpcIndex > 0 Then
                    TempCharIndex = MapData(map, X, Y + 1).NpcIndex
                    FoundChar = 2
                End If
            End If
            '?Es un personaje?
            If FoundChar = 0 Then
                If MapData(map, X, Y).Userindex > 0 Then
                    TempCharIndex = MapData(map, X, Y).Userindex
                    FoundChar = 1
                End If
                If MapData(map, X, Y).NpcIndex > 0 Then
                    TempCharIndex = MapData(map, X, Y).NpcIndex
                    FoundChar = 2
                End If
            End If
        End With
    
    
        'Reaccion al personaje
        If FoundChar = 1 Then '  ?Encontro un Usuario?
           If UserList(TempCharIndex).flags.AdminInvisible = 0 Or .flags.Privilegios And PlayerType.Dios Then
                With UserList(TempCharIndex)
                    If LenB(.DescRM) = 0 And .showName Then 'No tiene descRM y quiere que se vea su nombre.
                        If EsNewbie(TempCharIndex) Then
                            Stat = " <NEWBIE>"
                        End If
                        
                        If .Faccion.ArmadaReal = 1 Then
                            Stat = Stat & " <Ej?rcito Real> " & "<" & TituloReal(TempCharIndex) & ">"
                        ElseIf .Faccion.FuerzasCaos = 1 Then
                            Stat = Stat & " <Legi?n Oscura> " & "<" & TituloCaos(TempCharIndex) & ">"
                        End If
                        
                        If .GuildIndex > 0 Then
                            Stat = Stat & " <" & modGuilds.GuildName(.GuildIndex) & ">"
                        End If
                        
                        If Len(.desc) > 0 Then
                            Stat = "Ves a " & .Name & Stat & " - " & .desc
                        Else
                            Stat = "Ves a " & .Name & Stat
                        End If
                        
                                        
                        If .flags.Privilegios And PlayerType.RoyalCouncil Then
                            Stat = Stat & " [CONSEJO DE BANDERBILL]"
                            Ft = FontTypeNames.FONTTYPE_CONSEJOVesA
                        ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                            Stat = Stat & " [CONCILIO DE LAS SOMBRAS]"
                            Ft = FontTypeNames.FONTTYPE_EJECUCION
                        Else
                        If UserList(TempCharIndex).flags.Privilegios = PlayerType.Admin Then
                        Stat = Stat & " <Administrador> ~255~255~180~1~0"
                        ElseIf UserList(TempCharIndex).flags.Privilegios = PlayerType.Dios Then
                        Stat = Stat & " <Dios> ~250~250~150~1~0"
                        ElseIf UserList(TempCharIndex).flags.Privilegios = PlayerType.SemiDios Then
                        Stat = Stat & " <Semi-Dios> ~30~255~30~1~0"
                        ElseIf UserList(TempCharIndex).flags.Privilegios = PlayerType.Consejero Then
                        Stat = Stat & " <Consejero> ~30~150~30~1~0"
                                
                               ' Elijo el color segun el rango del GM:
                                ' Dios
                                If .flags.Privilegios = PlayerType.Dios Then
                                    Ft = FontTypeNames.fonttype_dios
                                ' Gm
                                ElseIf .flags.Privilegios = PlayerType.SemiDios Then
                                     Ft = FontTypeNames.FONTTYPE_GM
                                ' Conse
                                ElseIf .flags.Privilegios = PlayerType.Consejero Then
                                    Ft = FontTypeNames.FONTTYPE_CONSE
                                ' Rm o Dsrm
                                ElseIf .flags.Privilegios = (PlayerType.RoleMaster Or PlayerType.Consejero) Or .flags.Privilegios = (PlayerType.RoleMaster Or PlayerType.Dios) Then
                                    Ft = FontTypeNames.FONTTYPE_EJECUCION
                                End If
                                
                            ElseIf criminal(TempCharIndex) Then
                                Stat = Stat & " <CRIMINAL>"
                                Ft = FontTypeNames.FONTTYPE_FIGHT
                            Else
                                Stat = Stat & " <CIUDADANO>"
                                Ft = FontTypeNames.FONTTYPE_CITIZEN
                            End If
                        End If
                        
                        If .death = True Then
                            If .Pos.map = 195 Then
                            Stat = "<DeathMatch Player>"
                            Ft = FontTypeNames.FONTTYPE_CENTINELA
                        End If
                        End If
                        
                        If .hungry = True Then
                            If .Pos.map = 192 Then
                            Stat = "<JDH Player>"
                            Ft = FontTypeNames.FONTTYPE_CENTINELA
                        End If
                        End If
                        
                         If .flags.Premium > 0 Then
                        Stat = Stat & " [PREMIUM]"
                        Ft = FontTypeNames.FONTTYPE_PREMIUM
                        End If
                        
                        If StrComp(GreatPower.CurrentUser, UCase$(.Name)) = 0 Then
                            Stat = Stat & " <Gran poder>"
                        End If
                        
                        'esto envia al Ser vip el nombre <VIP> al clickear
                        
                        If .flags.DiosTerrenal > 0 Then
                            Stat = Stat & " [Dios Terrenal]"
                            Ft = FontTypeNames.FONTTYPE_TALK
                        End If
                        
If .flags.Oro > 0 Then
                            Stat = Stat & " [ORO]"
                            Ft = FontTypeNames.FONTTYPE_ORO
                        End If
                        
                        If .flags.Plata > 0 Then
                            Stat = Stat & " [PLATA]"
                            Ft = FontTypeNames.FONTTYPE_PLATA
                        End If
                        
                        If .flags.Bronce > 0 Then
                            Stat = Stat & " [BRONCE]"
                            Ft = FontTypeNames.FONTTYPE_BRONCE
                        End If
                        
            If .flags.Infectado > 0 Then
                            Stat = Stat & " [INFECTADO]"
                            Ft = FontTypeNames.FONTTYPE_CONSE
                        End If
                        
                                    If .flags.Angel > 0 Then
                            Stat = Stat & " [?NGEL]"
                            Ft = FontTypeNames.FONTTYPE_CONSEJOVesA
                        End If
                        
                                    If .flags.Demonio > 0 Then
                            Stat = Stat & " [DEMONIO]"
                            Ft = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
                        End If
                        
                        If .flags.Muerto = 1 Then
                               Stat = Stat & " <MUERTO>"
                               Ft = FontTypeNames.FONTTYPE_EJECUCION
                            End If
                    Else  'Si tiene descRM la muestro siempre.
                        Stat = .DescRM
                        Ft = FontTypeNames.FONTTYPE_INFOBOLD
                    End If
                End With
                
                If LenB(Stat) > 0 Then
                    Call WriteConsoleMsg(Userindex, Stat, Ft)
                End If
                
                FoundSomething = 1
                .flags.TargetUser = TempCharIndex
                .flags.TargetNPC = 0
                .flags.TargetNpcTipo = eNPCType.Comun
           End If
        End If
    
        With .flags
            If FoundChar = 2 Then '?Encontro un NPC?
                Dim estatus As String
                Dim MinHp As Long
                Dim MaxHp As Long
                Dim SupervivenciaSkill As Byte
                Dim sDesc As String
                
                MinHp = Npclist(TempCharIndex).Stats.MinHp
                MaxHp = Npclist(TempCharIndex).Stats.MaxHp
                SupervivenciaSkill = UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia)
                
                If .Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then
                    estatus = "(" & MinHp & "/" & MaxHp & ") "
                Else
                     If UserList(Userindex).flags.Muerto = 0 Then
                    If UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 10 Then
                        estatus = "(Dudoso) "
                    ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) > 10 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 20 Then
                        If Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp / 2) Then
                            estatus = "(Herido) "
                        Else
                            estatus = "(Sano) "
                        End If
                    ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) > 20 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 30 Then
                        If Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.5) Then
                            estatus = "(Malherido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.75) Then
                            estatus = "(Herido) "
                        Else
                            estatus = "(Sano) "
                        End If
                    ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) > 30 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 40 Then
                        If Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.25) Then
                            estatus = "(Muy malherido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.5) Then
                            estatus = "(Herido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.75) Then
                            estatus = "(Levemente herido) "
                        Else
                            estatus = "(Sano) "
                        End If
                    ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) > 40 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) < 60 Then
                        If Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.05) Then
                            estatus = "(Agonizando) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.1) Then
                            estatus = "(Casi muerto) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.25) Then
                            estatus = "(Muy Malherido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.5) Then
                            estatus = "(Herido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.75) Then
                            estatus = "(Levemente herido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp) Then
                            estatus = "(Sano) "
                        Else
                            estatus = "(Intacto) "
                        End If
                    ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 60 Then
                        estatus = "(" & Npclist(TempCharIndex).Stats.MinHp & "/" & Npclist(TempCharIndex).Stats.MaxHp & ") "
                    Else
                        estatus = "!error!"
                        End If
                    End If
                End If
                
                If Len(Npclist(TempCharIndex).desc) > 1 Then
                    Call WriteChatOverHead(Userindex, Npclist(TempCharIndex).desc, Npclist(TempCharIndex).Char.CharIndex, vbWhite)
                ElseIf TempCharIndex = CentinelaNPCIndex Then
                    'Enviamos nuevamente el texto del centinela seg?n quien pregunta
                    Call modCentinela.CentinelaSendClave(Userindex)
                Else
                    If Npclist(TempCharIndex).MaestroUser > 0 Then
                        Call WriteConsoleMsg(Userindex, estatus & Npclist(TempCharIndex).Name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).Name & ".", FontTypeNames.FONTTYPE_INFO)
                    Else
                        sDesc = estatus & Npclist(TempCharIndex).Name
                        If Npclist(TempCharIndex).Owner > 0 Then sDesc = sDesc & " le pertenece a " & UserList(Npclist(TempCharIndex).Owner).Name
                        sDesc = sDesc & "."
                        
                        Call WriteConsoleMsg(Userindex, sDesc, FontTypeNames.FONTTYPE_INFO)
                        
                        If .Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                            Call WriteConsoleMsg(Userindex, "Le peg? primero: " & Npclist(TempCharIndex).flags.AttackedFirstBy & ".", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                End If
                
                FoundSomething = 1
                .TargetNpcTipo = Npclist(TempCharIndex).NPCtype
                .TargetNPC = TempCharIndex
                .TargetUser = 0
                .TargetObj = 0
            End If
            
            If FoundChar = 0 Then
                .TargetNPC = 0
                .TargetNpcTipo = eNPCType.Comun
                .TargetUser = 0
            End If
            
            '*** NO ENCOTRO NADA ***
            If FoundSomething = 0 Then
                .TargetNPC = 0
                .TargetNpcTipo = eNPCType.Comun
                .TargetUser = 0
                .TargetObj = 0
                .TargetObjMap = 0
                .TargetObjX = 0
                .TargetObjY = 0
                Call WriteMultiMessage(Userindex, eMessages.DontSeeAnything)
            End If
        End With
    Else
        If FoundSomething = 0 Then
            With .flags
                .TargetNPC = 0
                .TargetNpcTipo = eNPCType.Comun
                .TargetUser = 0
                .TargetObj = 0
                .TargetObjMap = 0
                .TargetObjX = 0
                .TargetObjY = 0
            End With
            
            Call WriteMultiMessage(Userindex, eMessages.DontSeeAnything)
        End If
    End If
End With

Exit Sub

Errhandler:
    Call LogError("Error en LookAtTile. Error " & Err.Number & " : " & Err.Description)

End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As eHeading
'***************************************************
'Author: Unknown
'Last Modification: -
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'*****************************************************************

    Dim X As Integer
    Dim Y As Integer
    
    X = Pos.X - Target.X
    Y = Pos.Y - Target.Y
    
    'NE
    If Sgn(X) = -1 And Sgn(Y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.NORTH, eHeading.EAST)
        Exit Function
    End If
    
    'NW
    If Sgn(X) = 1 And Sgn(Y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.NORTH)
        Exit Function
    End If
    
    'SW
    If Sgn(X) = 1 And Sgn(Y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.SOUTH)
        Exit Function
    End If
    
    'SE
    If Sgn(X) = -1 And Sgn(Y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.SOUTH, eHeading.EAST)
        Exit Function
    End If
    
    'Sur
    If Sgn(X) = 0 And Sgn(Y) = -1 Then
        FindDirection = eHeading.SOUTH
        Exit Function
    End If
    
    'norte
    If Sgn(X) = 0 And Sgn(Y) = 1 Then
        FindDirection = eHeading.NORTH
        Exit Function
    End If
    
    'oeste
    If Sgn(X) = 1 And Sgn(Y) = 0 Then
        FindDirection = eHeading.WEST
        Exit Function
    End If
    
    'este
    If Sgn(X) = -1 And Sgn(Y) = 0 Then
        FindDirection = eHeading.EAST
        Exit Function
    End If
    
    'misma
    If Sgn(X) = 0 And Sgn(Y) = 0 Then
        FindDirection = 0
        Exit Function
    End If

End Function

Public Function ItemNoEsDeMapa(ByVal index As Integer, ByVal bIsExit As Boolean) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With ObjData(index)
        ItemNoEsDeMapa = .OBJType <> eOBJType.otPuertas And _
                    .OBJType <> eOBJType.otForos And _
                    .OBJType <> eOBJType.otCarteles And _
                    .OBJType <> eOBJType.otarboles And _
                    .OBJType <> eOBJType.otYacimiento And _
                    .OBJType <> eOBJType.otTeleport And _
                    Not (.OBJType = eOBJType.otTeleport And bIsExit)
    
    End With

End Function

Public Function MostrarCantidad(ByVal index As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With ObjData(index)
        MostrarCantidad = .OBJType <> eOBJType.otPuertas And _
                    .OBJType <> eOBJType.otForos And _
                    .OBJType <> eOBJType.otCarteles And _
                    .OBJType <> eOBJType.otarboles And _
                    .OBJType <> eOBJType.otYacimiento And _
                    .OBJType <> eOBJType.otTeleport
    End With

End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    EsObjetoFijo = OBJType = eOBJType.otForos Or _
                   OBJType = eOBJType.otCarteles Or _
                   OBJType = eOBJType.otarboles Or _
                   OBJType = eOBJType.otYacimiento
End Function
Private Function RhombLegalPos(ByRef Pos As WorldPos, ByRef vX As Long, ByRef vY As Long, _
                               ByVal Distance As Long, Optional PuedeAgua As Boolean = False, _
                               Optional PuedeTierra As Boolean = True, _
                               Optional ByVal CheckExitTile As Boolean = False) As Boolean
'***************************************************
'Author: Marco Vanotti (Marco)
'Last Modification: -
' walks all the perimeter of a rhomb of side  "distance + 1",
' which starts at Pos.x - Distance and Pos.y
'***************************************************

    Dim i As Long
    
    vX = Pos.X - Distance
    vY = Pos.Y
    
    For i = 0 To Distance - 1
        If (LegalPos(Pos.map, vX + i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX + i
            vY = vY - i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    vX = Pos.X
    vY = Pos.Y - Distance
    
    For i = 0 To Distance - 1
        If (LegalPos(Pos.map, vX + i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX + i
            vY = vY + i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    vX = Pos.X + Distance
    vY = Pos.Y
    
    For i = 0 To Distance - 1
        If (LegalPos(Pos.map, vX - i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX - i
            vY = vY + i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    vX = Pos.X
    vY = Pos.Y + Distance
    
    For i = 0 To Distance - 1
        If (LegalPos(Pos.map, vX - i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX - i
            vY = vY - i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    RhombLegalPos = False
    
End Function

Public Function RhombLegalTilePos(ByRef Pos As WorldPos, ByRef vX As Long, ByRef vY As Long, _
                                  ByVal Distance As Long, ByVal objindex As Integer, ByVal ObjAmount As Long, _
                                  ByVal PuedeAgua As Boolean, ByVal PuedeTierra As Boolean) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: -
' walks all the perimeter of a rhomb of side  "distance + 1",
' which starts at Pos.x - Distance and Pos.y
' and searchs for a valid position to drop items
'***************************************************
On Error GoTo Errhandler

    Dim i As Long
    Dim hayobj As Boolean
    
    Dim X As Integer
    Dim Y As Integer
    Dim MapObjIndex As Integer
    
    vX = Pos.X - Distance
    vY = Pos.Y
    
    For i = 0 To Distance - 1
        
        X = vX + i
        Y = vY - i
        
        If (LegalPos(Pos.map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.map, X, Y, objindex, ObjAmount) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True
                Exit Function
            End If
            
        End If
    Next
    
    vX = Pos.X
    vY = Pos.Y - Distance
    
    For i = 0 To Distance - 1
        
        X = vX + i
        Y = vY + i
        
        If (LegalPos(Pos.map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.map, X, Y, objindex, ObjAmount) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    
    vX = Pos.X + Distance
    vY = Pos.Y
    
    For i = 0 To Distance - 1
        
        X = vX - i
        Y = vY + i
    
        If (LegalPos(Pos.map, X, Y, PuedeAgua, PuedeTierra, True)) Then
        
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.map, X, Y, objindex, ObjAmount) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    
    vX = Pos.X
    vY = Pos.Y + Distance
    
    For i = 0 To Distance - 1
        
        X = vX - i
        Y = vY - i
    
        If (LegalPos(Pos.map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.map, X, Y, objindex, ObjAmount) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    
    RhombLegalTilePos = False
    
    Exit Function
    
Errhandler:
    Call LogError("Error en RhombLegalTilePos. Error: " & Err.Number & " - " & Err.Description)
End Function

Public Function HayObjeto(ByVal Mapa As Integer, ByVal X As Long, ByVal Y As Long, _
                          ByVal objindex As Integer, ByVal ObjAmount As Long) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: -
'Checks if there's space in a tile to add an itemAmount
'***************************************************
    Dim MapObjIndex As Integer
    MapObjIndex = MapData(Mapa, X, Y).ObjInfo.objindex
            
    ' Hay un objeto tirado?
    If MapObjIndex <> 0 Then
        ' Es el mismo objeto?
        If MapObjIndex = objindex Then
            ' La suma es menor a 10k?
            HayObjeto = (MapData(Mapa, X, Y).ObjInfo.Amount + ObjAmount > MAX_INVENTORY_OBJS)
        Else
            HayObjeto = True
        End If
    Else
        HayObjeto = False
    End If

End Function
Public Function EsPremium(ByVal Userindex As Integer) As Boolean 'es un usuario premium?
EsPremium = UserList(Userindex).flags.Premium
End Function

Public Sub ShowMenu(ByVal Userindex As Integer, ByVal map As Integer, _
    ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 10/05/2010
'Shows menu according to user, npc or object right clicked.
'***************************************************

On Error GoTo Errhandler

    With UserList(Userindex)
        
        ' In Vision Range
        If (Abs(.Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(.Pos.X - X) > RANGO_VISION_X) Then Exit Sub
        
        ' Valid position?
        If Not InMapBounds(map, X, Y) Then Exit Sub
        
        With .flags
            ' Alive?
            If .Muerto = 1 Then Exit Sub
            
            ' Trading?
            If .Comerciando Then Exit Sub
            
            ' Reset flags
            .TargetNPC = 0
            .TargetNpcTipo = eNPCType.Comun
            .TargetUser = 0
            .TargetObj = 0
            .TargetObjMap = 0
            .TargetObjX = 0
            .TargetObjY = 0
            
            .TargetMap = map
            .TargetX = X
            .TargetY = Y
            
            Dim tmpIndex As Integer
            Dim FoundChar As Byte
            Dim MenuIndex As Integer
            
            ' Npc or user? (lower position)
            If Y + 1 <= YMaxMapSize Then
                
                ' User?
                tmpIndex = MapData(map, X, Y + 1).Userindex
                If tmpIndex > 0 Then
                    ' Invalid: Admin invisible, user invisible or hidden (if not is userindex)
                    If (UserList(tmpIndex).flags.AdminInvisible Or _
                        UserList(tmpIndex).flags.invisible Or _
                        UserList(tmpIndex).flags.Oculto) = 0 Or tmpIndex = Userindex Then
                        
                        FoundChar = 1
                    End If
                End If
                
                ' Npc?
                If MapData(map, X, Y + 1).NpcIndex > 0 Then
                    tmpIndex = MapData(map, X, Y + 1).NpcIndex
                    FoundChar = 2
                End If
            End If
             
            ' Npc or user? (upper position)
            If FoundChar = 0 Then
                
                ' User?
                tmpIndex = MapData(map, X, Y).Userindex
                If tmpIndex > 0 Then
                    ' Invalid: Admin invisible, user invisible or hidden (if not is userindex)
                    If (UserList(tmpIndex).flags.AdminInvisible Or _
                        UserList(tmpIndex).flags.invisible Or _
                        UserList(tmpIndex).flags.Oculto) = 0 Or tmpIndex = Userindex Then
                        
                        FoundChar = 1
                    End If
                End If
                
                ' Npc?
                If MapData(map, X, Y).NpcIndex > 0 Then
                    tmpIndex = MapData(map, X, Y).NpcIndex
                    FoundChar = 2
                End If
            End If
            
            ' User
            If FoundChar = 1 Then
                MenuIndex = eMenues.ieUser
                    
                .TargetUser = tmpIndex
                
            ' Npc
            ElseIf FoundChar = 2 Then
                ' Has menu attached?
                'If Npclist(tmpIndex).MenuIndex <> 0 Then
                    'MenuIndex = Npclist(tmpIndex).MenuIndex
               ' End If
                
                '.TargetNpcTipo = Npclist(tmpIndex).NPCtype
                '.TargetNPC = tmpIndex
            End If
            
            ' No user or npc found
            If FoundChar = 0 Then
                
                ' Is there any object?
                tmpIndex = MapData(map, X, Y).ObjInfo.objindex
                If tmpIndex > 0 Then
                    ' Has menu attached?
                    'MenuIndex = ObjData(tmpIndex).MenuIndex
                    
                    'If MenuIndex = eMenues.ieFogata Then
                        'If .Descansar = 1 Then MenuIndex = eMenues.ieFogataDescansando
                    'End If
                    
                    '.TargetObjMap = Map
                    '.TargetObjX = X
                   ' .TargetObjY = Y
                End If
            End If
        End With
    End With
    
    ' Show it
    If MenuIndex <> 0 Then _
        Call WriteShowMenu(Userindex, MenuIndex)
    
    Exit Sub

Errhandler:
    Call LogError("Error en ShowMenu. Error " & Err.Number & " : " & Err.Description)
End Sub

