Attribute VB_Name = "UsUaRiOs"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Public Sub ActStats(ByVal victimIndex As Integer, ByVal AttackerIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 11/03/2010
'11/03/2010: ZaMa - Ahora no te vuelve cirminal por matar un atacable
'***************************************************

    Dim DaExp As Integer
    Dim EraCriminal As Boolean
    
    DaExp = CInt(UserList(victimIndex).Stats.ELV) * 2
    
    With UserList(AttackerIndex)
        .Stats.Exp = .Stats.Exp + DaExp
        If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        
        If TriggerZonaPelea(victimIndex, AttackerIndex) <> TRIGGER6_PERMITE Then
        
            ' Es legal matarlo si estaba en atacable
            If UserList(victimIndex).flags.AtacablePor <> AttackerIndex Then
                EraCriminal = criminal(AttackerIndex)
                
                With .Reputacion
                    If Not criminal(victimIndex) Then
                        .AsesinoRep = .AsesinoRep + vlASESINO * 2
                        'If .AsesinoRep > MAXREP Then .AsesinoRep = MAXREP
                        .BurguesRep = 0
                        .NobleRep = 0
                        .PlebeRep = 0
                    Else
                        .NobleRep = .NobleRep + vlNoble
                        If .NobleRep > MAXREP Then .NobleRep = MAXREP
                    End If
                End With
                
                If criminal(AttackerIndex) Then
                    If Not EraCriminal Then Call RefreshCharStatus(AttackerIndex)
                Else
                    If EraCriminal Then Call RefreshCharStatus(AttackerIndex)
                End If
            End If
        End If
        
        
                If UserList(victimIndex).flags.Infectado = 1 Then
                    'A resetear variables...
                    UserList(victimIndex).Stats.MaxHp = UserList(victimIndex).Stats.ViejaHP
                    UserList(victimIndex).Stats.MaxMAN = UserList(victimIndex).Stats.ViejaMan
                    UserList(victimIndex).Stats.MinHp = UserList(victimIndex).Stats.ViejaminHP
                    UserList(victimIndex).Stats.MinMAN = UserList(victimIndex).Stats.ViejaminMan
                    UserList(victimIndex).flags.Infectado = 0
                    RefreshCharStatus victimIndex 'estoy seguro que el refresh lo usa más tarde, pero we
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡" & .Name & " ha derrotado al infectado!", FontTypeNames.FONTTYPE_CONSE))
                
                '////////////////////PREMIO\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim MiPremio As Obj
                Dim Oro As Long
                MiPremio.Amount = 1 'cantidad del premio..
                MiPremio.objindex = 402 ' objeto premio según obj.dat
                Oro = 300000 ' premio = 1kk ...
                .Stats.Gld = .Stats.Gld + Oro
                Call MeterItemEnInventario(AttackerIndex, MiPremio)
                End If

        If UserList(victimIndex).flags.Angel = 1 Then
'A resetear variables...
UserList(victimIndex).Stats.MaxHp = UserList(victimIndex).Stats.ViejaHP
UserList(victimIndex).Stats.MaxMAN = UserList(victimIndex).Stats.ViejaMan
UserList(victimIndex).Stats.MinHp = UserList(victimIndex).Stats.ViejaminHP
UserList(victimIndex).Stats.MinMAN = UserList(victimIndex).Stats.ViejaminMan
UserList(victimIndex).flags.Angel = 0
RefreshCharStatus victimIndex 'estoy seguro que el refresh lo usa más tarde, pero we
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡" & .Name & " ha derrotado al Ángel!", FontTypeNames.FONTTYPE_EJECUCION))
'////////////////////PREMIO\\\\\\\\\\\\\\\\\\\\\\\\\\
MiPremio.Amount = 1 'cantidad del premio..
MiPremio.objindex = 402 ' objeto premio según obj.dat
Oro = 300000 ' premio = 1kk ...
.Stats.Gld = .Stats.Gld + Oro
Call MeterItemEnInventario(AttackerIndex, MiPremio)
End If

        If UserList(victimIndex).flags.Demonio = 1 Then
'A resetear variables...
UserList(victimIndex).Stats.MaxHp = UserList(victimIndex).Stats.ViejaHP
UserList(victimIndex).Stats.MaxMAN = UserList(victimIndex).Stats.ViejaMan
UserList(victimIndex).Stats.MinHp = UserList(victimIndex).Stats.ViejaminHP
UserList(victimIndex).Stats.MinMAN = UserList(victimIndex).Stats.ViejaminMan
UserList(victimIndex).flags.Demonio = 0
RefreshCharStatus victimIndex 'estoy seguro que el refresh lo usa más tarde, pero we
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡" & .Name & " ha derrotado al Demonio!", FontTypeNames.FONTTYPE_EJECUCION))
'////////////////////PREMIO\\\\\\\\\\\\\\\\\\\\\\\\\\
MiPremio.Amount = 1 'cantidad del premio..
MiPremio.objindex = 402 ' objeto premio según obj.dat
Oro = 300000 ' premio = 1kk ...
.Stats.Gld = .Stats.Gld + Oro
Call MeterItemEnInventario(AttackerIndex, MiPremio)
End If
        
        
            'Eduardo
    If UserList(victimIndex).Name = ElmasbuscadoFusion Then
    Dim Recom As Obj
    Recom.Amount = 1
    Recom.objindex = 402 'Aka si quieren reemplazan por el items suyos a ganar :D
    Call MeterItemEnInventario(AttackerIndex, Recom)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡¡Atencion!!: " & UserList(victimIndex).Name & " ha asesinado a " & UserList(victimIndex).Name & " (el mas buscado).", FontTypeNames.FONTTYPE_GUILD))
    Call WriteConsoleMsg(AttackerIndex, "Has matado al él más buscado. La recompensa ha sido entregada en tu inventario.", FontTypeNames.FONTTYPE_GUILD)
    ElmasbuscadoFusion = 0
    End If
        
        'Lo mata
        'Call WriteConsoleMsg(attackerIndex, "Has matado a " & UserList(VictimIndex).name & "!", FontTypeNames.FONTTYPE_FIGHT)
        'Call WriteConsoleMsg(attackerIndex, "Has ganado " & DaExp & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
        'Call WriteConsoleMsg(VictimIndex, "¡" & .name & " te ha matado!", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteMultiMessage(AttackerIndex, eMessages.HaveKilledUser, victimIndex, DaExp)
        Call WriteMultiMessage(victimIndex, eMessages.UserKill, AttackerIndex)

        'Call UserDie(VictimIndex)
        Call FlushBuffer(victimIndex)
        
        'Log
        Call LogAsesinato(.Name & " asesino a " & UserList(victimIndex).Name)
    End With
End Sub
Public Sub RevivirUsuario(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)
        .flags.Muerto = 0
        .Stats.MinHp = .Stats.UserAtributos(eAtributos.Constitucion)
        
        If .Stats.MinHp > .Stats.MaxHp Then
            .Stats.MinHp = .Stats.MaxHp
        End If
        
        If .flags.Navegando = 1 Then
            Call ToggleBoatBody(UserIndex)
        Else
            Call DarCuerpoDesnudo(UserIndex)
            
            .Char.Head = .OrigChar.Head
        End If
        
        If .flags.Traveling Then
            .flags.Traveling = 0
            .Counters.goHome = 0
            Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
        End If
        
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        Call WriteUpdateUserStats(UserIndex)
    End With
End Sub



Public Sub ChangeUserChar(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal casco As Integer, Optional ByVal Transformation As Boolean = False)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    With UserList(UserIndex).Char
        .body = body
        .Head = Head
        .Heading = Heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = casco
        
        
        ' En caso de que recien transforme al usuario cambiamos de body.
        If Not Transformation Then
            If UserList(UserIndex).flags.SlotEvent > 0 Then
                If Events(UserList(UserIndex).flags.SlotEvent).CharBody <> 0 Then
                    Exit Sub
                End If
            End If
        End If
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(body, Head, Heading, .CharIndex, Arma, Escudo, .FX, .loops, casco))
    End With
End Sub

Public Function GetWeaponAnim(ByVal UserIndex As Integer, ByVal objindex As Integer) As Integer
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 03/29/10
'
'***************************************************
    Dim tmp As Integer

    With UserList(UserIndex)
        tmp = ObjData(objindex).WeaponRazaEnanaAnim
            
        If tmp > 0 Then
            If .raza = eRaza.Enano Or .raza = eRaza.Gnomo Then
                GetWeaponAnim = tmp
                Exit Function
            End If
        End If
        
        GetWeaponAnim = ObjData(objindex).WeaponAnim
    End With
End Function

Public Sub EnviarFama(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim L As Long
    
    With UserList(UserIndex).Reputacion
        L = (-.AsesinoRep) + _
            (-.BandidoRep) + _
            .BurguesRep + _
            (-.LadronesRep) + _
            .NobleRep + _
            .PlebeRep
        L = Round(L / 6)
        
        .Promedio = L
    End With
    
    Call WriteFame(UserIndex)
End Sub

Public Sub EraseUserChar(ByVal UserIndex As Integer, ByVal IsAdminInvisible As Boolean)
'*************************************************
'Author: Unknown
'Last modified: 08/01/2009
'08/01/2009: ZaMa - No se borra el char de un admin invisible en todos los clientes excepto en su mismo cliente.
'*************************************************

On Error GoTo ErrorHandler
    
    With UserList(UserIndex)
        CharList(.Char.CharIndex) = 0
        
        If .Char.CharIndex = LastChar Then
            Do Until CharList(LastChar) > 0
                LastChar = LastChar - 1
                If LastChar <= 1 Then Exit Do
            Loop
        End If
        
        ' Si esta invisible, solo el sabe de su propia existencia, es innecesario borrarlo en los demas clientes
        If IsAdminInvisible Then
            Call EnviarDatosASlot(UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
        Else
            'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
        End If
        
        Call QuitarUser(UserIndex, .Pos.map)
        
        MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex = 0
        .Char.CharIndex = 0
    End With
    
    NumChars = NumChars - 1
Exit Sub
    
ErrorHandler:
    Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.Description)
End Sub

Public Sub RefreshCharStatus(ByVal UserIndex As Integer)
'*************************************************
'Author: Tararira
'Last modified: 04/07/2009
'Refreshes the status and tag of UserIndex.
'04/07/2009: ZaMa - Ahora mantenes la fragata fantasmal si estas muerto.
'*************************************************
    Dim ClanTag As String
    Dim NickColor As Byte
    
    With UserList(UserIndex)
        If .GuildIndex > 0 Then
            ClanTag = modGuilds.GuildName(.GuildIndex)
            ClanTag = " <" & ClanTag & ">"
        End If
        
        NickColor = GetNickColor(UserIndex)
        
        If .showName Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, NickColor, .Name & ClanTag))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, NickColor, vbNullString))
        End If
        
         If .flags.Infectado = 1 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, NickColor, .Name & ClanTag, 1))
        End If
        
                 If .flags.Angel = 1 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, NickColor, .Name & ClanTag, 1))
        End If
        
                 If .flags.Demonio = 1 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, NickColor, .Name & ClanTag, 1))
        End If
        
       'Si esta navengando, se cambia la barca.
        If .flags.Navegando Then
            If .flags.Muerto = 1 Then
                .Char.body = iFragataFantasmal
            Else
                Call ToggleBoatBody(UserIndex)
            End If
            
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
    End With
End Sub

Public Function GetNickColor(ByVal UserIndex As Integer) As Byte
'*************************************************
'Author: ZaMa
'Last modified: 15/01/2010
'
'*************************************************
    
    With UserList(UserIndex)
        
        If criminal(UserIndex) Then
            GetNickColor = eNickColor.ieCriminal
        Else
            GetNickColor = eNickColor.ieCiudadano
        End If
        
        If .flags.AtacablePor > 0 Then GetNickColor = GetNickColor Or eNickColor.ieAtacable
        
        If .flags.SlotEvent > 0 Then
            With Events(.flags.SlotEvent)
                If .Modality = CastleMode Then
                    If .Users(UserList(UserIndex).flags.SlotUserEvent).Team = 1 Then
                        GetNickColor = eNickColor.ieTeamUno
                    ElseIf .Users(UserList(UserIndex).flags.SlotUserEvent).Team = 2 Then
                        GetNickColor = eNickColor.ieTeamDos
                    End If
                    End If
                    If .Modality = Enfrentamientos Then
                    If .Users(UserList(UserIndex).flags.SlotUserEvent).Team = 1 Then
                        GetNickColor = eNickColor.ieTeamUno
                    ElseIf .Users(UserList(UserIndex).flags.SlotUserEvent).Team = 2 Then
                        GetNickColor = eNickColor.ieTeamDos
                    End If
                End If
            End With
        End If
    End With
    
End Function

Public Sub MakeUserChar(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal UserIndex As Integer, _
        ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ButIndex As Boolean = False)
'*************************************************
'Author: Unknown
'Last modified: 15/01/2010
'23/07/2009: Budi - Ahora se envía el nick
'15/01/2010: ZaMa - Ahora se envia el color del nick.
'*************************************************

On Error GoTo Errhandler

    Dim CharIndex As Integer
    Dim ClanTag As String
    Dim NickColor As Byte
    Dim UserName As String
    Dim Privileges As Byte
    
    With UserList(UserIndex)
    
        If InMapBounds(map, X, Y) Then
            'If needed make a new character in list
            If .Char.CharIndex = 0 Then
                CharIndex = NextOpenCharIndex
                .Char.CharIndex = CharIndex
                CharList(CharIndex) = UserIndex
            End If
            
            'Place character on map if needed
            If toMap Then MapData(map, X, Y).UserIndex = UserIndex
            
            'Send make character command to clients
            If Not toMap Then
                If .GuildIndex > 0 Then
                    ClanTag = modGuilds.GuildName(.GuildIndex)
                End If
                
                NickColor = GetNickColor(UserIndex)
                Privileges = .flags.Privilegios
                
                'Preparo el nick
                If .showName Then
                    UserName = .Name
                    
                    If .flags.EnConsulta Then
                        UserName = UserName & " " & TAG_CONSULT_MODE
                    Else
                        If UserList(sndIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
                            If LenB(ClanTag) <> 0 Then _
                                UserName = UserName & " <" & ClanTag & ">"
                        Else
                            If (.flags.invisible Or .flags.Oculto) And (Not .flags.AdminInvisible = 1) Then
                                UserName = UserName & " " & TAG_USER_INVISIBLE
                            Else
                                If LenB(ClanTag) <> 0 Then _
                                    UserName = UserName & " <" & ClanTag & ">"
                            End If
                        End If
                    End If
                End If
            
                Call WriteCharacterCreate(sndIndex, .Char.body, .Char.Head, .Char.Heading, _
                            .Char.CharIndex, X, Y, _
                            .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, 999, .Char.CascoAnim, _
                            UserName, NickColor, Privileges)
            Else
                'Hide the name and clan - set privs as normal user
                 Call AgregarUser(UserIndex, .Pos.map, ButIndex)
            End If
        End If
    End With
Exit Sub

Errhandler:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.Description)
    'Resume Next
    Call CloseSocket(UserIndex)
End Sub

''
' Checks if the user gets the next level.
'
' @param UserIndex Specifies reference to user

Public Sub CheckUserLevel(ByVal UserIndex As Integer)
    On Error GoTo CheckUserLevel_Error
      '*************************************************
      'Author: Unknown
      'Last modified: 11/19/2009
      'Chequea que el usuario no halla alcanzado el siguiente nivel,
      'de lo contrario le da la vida, mana, etc, correspodiente.
      '07/08/2006 Integer - Modificacion de los valores
      '01/10/2007 Tavo - Corregido el BUG de STAT_MAXELV

      'kUserLevel - Error : 70 - Description : Permission denied
      'Tenes que abrir el Server ne el vps como admin...
      '24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones en ELU al subir de nivel.
      '24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones de la subida de mana de los magos por lvl.
      '13/03/2007 Pablo (ToxicWaste) - Agrego diferencias entre el 18 y el 19 en Constitución.
      '09/01/2008 Pablo (ToxicWaste) - Ahora el incremento de vida por Consitución se controla desde Balance.dat
      '12/09/2008 Marco Vanotti (Marco) - Ahora si se llega a nivel 25 y está en un clan, se lo expulsa para no sumar antifacción
      '02/03/2009 ZaMa - Arreglada la validacion de expulsion para miembros de clanes faccionarios que llegan a 25.
      '11/19/2009 Pato - Modifico la nueva fórmula de maná ganada para el bandido y se la limito a 499
      '02/04/2010: ZaMa - Modifico la ganancia de hit por nivel del ladron.
      '*************************************************
          Dim Pts As Integer
          Dim AumentoHIT As Integer
          Dim AumentoMANA As Integer
          Dim AumentoSTA As Integer
          Dim AumentoHP As Integer
          Dim WasNewbie As Boolean
             Dim WasQuince As Boolean
             Dim WasSiete As Boolean
             Dim WasOcho As Boolean
             Dim WasNueve As Boolean
             Dim WasQuinceElv As Boolean
             Dim WasVeinte As Boolean
             Dim WasVeinticinco As Boolean
             Dim WasQuinceM As Boolean
             Dim WasTreintaM As Boolean
             Dim WasHM As Boolean
             Dim WasUM As Boolean
             Dim WasMM As Boolean
             Dim WasVip As Boolean
             Dim WasVipp As Boolean
             Dim WasVipb As Boolean
             Dim WasNoUM As Boolean
      Dim waspremium As Boolean
          Dim Promedio As Double
          Dim aux As Integer
          Dim DistVida(1 To 5) As Integer
          Dim GI As Integer 'Guild Index
          
10        WasNewbie = EsNewbie(UserIndex)
20        WasQuince = EsQuince(UserIndex)
30        WasSiete = EsSiete(UserIndex)
40        WasOcho = EsOcho(UserIndex)
50        WasNueve = EsNueve(UserIndex)
60        WasQuince = EsQuince(UserIndex)
70        WasVeinte = EsVeinte(UserIndex)
80        WasVeinticinco = EsVeinticinco(UserIndex)
90        WasQuinceM = EsQuinceM(UserIndex)
100       WasTreintaM = EsTreintaM(UserIndex)
110       WasHM = EsHM(UserIndex)
120       WasUM = EsUM(UserIndex)
130       WasMM = EsMM(UserIndex)
140       WasVip = EsVip(UserIndex)
150       WasVipp = EsVipp(UserIndex)
160       WasVipb = EsVipb(UserIndex)
170       WasNoUM = NoEsUM(UserIndex)
180       waspremium = EsPremium(UserIndex)
          
190       With UserList(UserIndex)
200           Do While .Stats.Exp >= .Stats.ELU
                  
                  'Checkea si alcanzó el máximo nivel
210               If .Stats.ELV >= STAT_MAXELV Then
220                   .Stats.Exp = 0
230                   .Stats.ELU = 0
240                   Exit Sub
250               End If
                  
                  'Store it!
260               Call Statistics.UserLevelUp(UserIndex)
                  
                  If .Stats.ELV >= 48 Then
270               Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, .Pos.X, .Pos.Y))
280               Call WriteConsoleMsg(UserIndex, "¡Has subido de nivel!", FontTypeNames.FONTTYPE_INFO)
                End If
                  
290               If .Stats.ELV = 1 Then
300                   Pts = 0
310               Else
                      'For multiple levels being rised at once
320                   Pts = Pts + 0
330               End If
                  
340               .Stats.ELV = .Stats.ELV + 1
                  
350               .Stats.Exp = .Stats.Exp - .Stats.ELU
                  
                  'Nueva subida de exp x lvl. Pablo (ToxicWaste)
360               If .Stats.ELV = 2 Then
370                   .Stats.ELU = 450
380                   ElseIf .Stats.ELV = 3 Then
390                   .Stats.ELU = 675
400                   ElseIf .Stats.ELV = 4 Then
410                   .Stats.ELU = 1012
420                   ElseIf .Stats.ELV = 5 Then
430                   .Stats.ELU = 1518
440                   ElseIf .Stats.ELV = 6 Then
450                   .Stats.ELU = 2277
460                   ElseIf .Stats.ELV = 7 Then
470                   .Stats.ELU = 3416
480                   ElseIf .Stats.ELV = 8 Then
490                   .Stats.ELU = 5124
500                   ElseIf .Stats.ELV = 9 Then
510                   .Stats.ELU = 7886
520                   ElseIf .Stats.ELV = 10 Then
530                   .Stats.ELU = 11529
540                   ElseIf .Stats.ELV = 11 Then
550                   .Stats.ELU = 14988
560                   ElseIf .Stats.ELV = 12 Then
570                   .Stats.ELU = 19484
580                   ElseIf .Stats.ELV = 13 Then
590                   .Stats.ELU = 25329
600                   ElseIf .Stats.ELV = 14 Then
610                   .Stats.ELU = 32928
620                   ElseIf .Stats.ELV = 15 Then
630                   .Stats.ELU = 42806
640                   ElseIf .Stats.ELV = 16 Then
650                   .Stats.ELU = 55648
660                   ElseIf .Stats.ELV = 17 Then
670                   .Stats.ELU = 72342
680                   ElseIf .Stats.ELV = 18 Then
690                   .Stats.ELU = 94045
700                   ElseIf .Stats.ELV = 19 Then
710                   .Stats.ELU = 122259
720                   ElseIf .Stats.ELV = 20 Then
730                   .Stats.ELU = 158937
740                   ElseIf .Stats.ELV = 21 Then
750                   .Stats.ELU = 206618
760                   ElseIf .Stats.ELV = 22 Then
770                   .Stats.ELU = 268603
780                   ElseIf .Stats.ELV = 23 Then
790                   .Stats.ELU = 349184
800                   ElseIf .Stats.ELV = 24 Then
810                   .Stats.ELU = 453939
820                   ElseIf .Stats.ELV = 25 Then
830                   .Stats.ELU = 544727
840                   ElseIf .Stats.ELV = 26 Then
850                   .Stats.ELU = 667632
860                   ElseIf .Stats.ELV = 27 Then
870                   .Stats.ELU = 784406
880                   ElseIf .Stats.ELV = 28 Then
890                   .Stats.ELU = 941287
900                   ElseIf .Stats.ELV = 29 Then
910                   .Stats.ELU = 1129544
920                   ElseIf .Stats.ELV = 30 Then
930                   .Stats.ELU = 1355453
940                   ElseIf .Stats.ELV = 31 Then
950                   .Stats.ELU = 1626544
960                   ElseIf .Stats.ELV = 32 Then
970                   .Stats.ELU = 1951853
980                   ElseIf .Stats.ELV = 33 Then
990                   .Stats.ELU = 2342224
1000                  ElseIf .Stats.ELV = 34 Then
1010                  .Stats.ELU = 3372803
1020                  ElseIf .Stats.ELV = 35 Then
1030                  .Stats.ELU = 4047364
1040                  ElseIf .Stats.ELV = 36 Then
1050                  .Stats.ELU = 5828204
1060                  ElseIf .Stats.ELV = 37 Then
1070                  .Stats.ELU = 6993845
1080                  ElseIf .Stats.ELV = 38 Then
1090                  .Stats.ELU = 8392614
1100                  ElseIf .Stats.ELV = 39 Then
1110                  .Stats.ELU = 10071137
1120                  ElseIf .Stats.ELV = 40 Then
1130                  .Stats.ELU = 120853640
1140                  ElseIf .Stats.ELV = 41 Then
1150                  .Stats.ELU = 145024370
1160                  ElseIf .Stats.ELV = 42 Then
1170                  .Stats.ELU = 174029240
1180                  ElseIf .Stats.ELV = 43 Then
1190                  .Stats.ELU = 208835090
1200                  ElseIf .Stats.ELV = 44 Then
1210                  .Stats.ELU = 417670180
1220                  ElseIf .Stats.ELV = 45 Then
1230                  .Stats.ELU = 835340360
1240                  ElseIf .Stats.ELV = 46 Then
1250                  .Stats.ELU = 1670680720
1260                  Else
1270                  .Stats.ELU = 0
1280                  End If
                  
1290         Select Case .clase
                 Case eClass.Warrior
1300                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
1310                      AumentoHP = RandomNumber(9, 12)
1320                      Case 20
1330                      AumentoHP = RandomNumber(8, 12)
1340                      Case 19
1350                      AumentoHP = RandomNumber(8, 11)
1360                      Case 18
1370                      AumentoHP = RandomNumber(7, 11)
1380                      Case Else
1390                      AumentoHP = RandomNumber(6, .Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPGuerrero
1400                      End Select
                          
                          
1410                      If (.Stats.ELV < 48) Then
1420                      AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
1430                      Else
1440                      AumentoHIT = 1
1450                      End If
                          
1460                      AumentoSTA = AumentoSTDef
                      
1470                  Case eClass.Hunter
1480                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
1490                      AumentoHP = RandomNumber(9, 11)
1500                      Case 20
1510                      AumentoHP = RandomNumber(8, 11)
1520                      Case 19
1530                      AumentoHP = RandomNumber(7, 11)
1540                      Case 18
1550                      AumentoHP = RandomNumber(6, 10)
1560                      Case Else
1570                      AumentoHP = RandomNumber(6, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)
1580                      End Select
                      
                      
1590                  If (.Stats.ELV < 48) Then
1600                      AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
1610                      Else
1620                      AumentoHIT = 1
1630                      End If
                          
1640                      AumentoSTA = AumentoSTDef
                      
1650                  Case eClass.Pirat
1660                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
1670                      AumentoHP = RandomNumber(9, 11)
1680                      Case 20
1690                      AumentoHP = RandomNumber(8, 11)
1700                      Case 19
1710                      AumentoHP = RandomNumber(7, 11)
1720                      Case 18
1730                      AumentoHP = RandomNumber(6, 11)
1740                      Case Else
1750                      AumentoHP = RandomNumber(4, .Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPGuerrero
1760                      End Select
                      
1770                  If (.Stats.ELV < 48) Then
1780                      AumentoHIT = 3
1790                      Else
1800                      AumentoHIT = 2
1810                      End If
                          
1820                      AumentoSTA = AumentoSTDef
                      
1830                  Case eClass.Paladin
1840                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
1850                      AumentoHP = RandomNumber(9, 11)
1860                      Case 20
1870                      AumentoHP = RandomNumber(8, 11)
1880                      Case 19
1890                      AumentoHP = RandomNumber(7, 11)
1900                      Case 18
1910                      AumentoHP = RandomNumber(6, 11)
1920                      Case Else
1930                      AumentoHP = RandomNumber(4, .Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPCazador
1940                      End Select
                      
                  
1950              If (.Stats.ELV > 47) Then
1960              AumentoHP = .Stats.UserAtributos(eAtributos.Constitucion) \ 4 + AdicionalHPCazador
1970              End If
                  
1980                 If (.Stats.ELV < 48) Then
1990                      AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
2000                      Else
2010                      AumentoHIT = 1
2020                      End If
                          
2030                If (.Stats.ELV < 48) Then
2040                AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
2050                Else
2060                AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia) \ 2
2070                End If
                    
2080                      AumentoSTA = AumentoSTDef
                      
2090                  Case eClass.Thief
2100                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
2110                      AumentoHP = RandomNumber(6, 9)
2120                      Case 20
2130                      AumentoHP = RandomNumber(5, 9)
2140                      Case 19
2150                      AumentoHP = RandomNumber(4, 9)
2160                      Case 18
2170                      AumentoHP = RandomNumber(4, 8)
2180                      Case Else
2190                      AumentoHP = RandomNumber(4, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)
2200                      End Select
                      
2210                      AumentoHIT = 2
2220                      AumentoSTA = AumentoSTLadron
                      
2230                  Case eClass.Mage
2240                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
2250                      AumentoHP = RandomNumber(6, 9)
2260                      Case 20
2270                      AumentoHP = RandomNumber(5, 8)
2280                      Case 19
2290                      AumentoHP = RandomNumber(4, 8)
2300                      Case 18
2310                      AumentoHP = RandomNumber(3, 8)
2320                      Case Else
2330                      AumentoHP = RandomNumber(5, .Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
2340                      End Select
2350                      If AumentoHP < 1 Then AumentoHP = 4
                          
2360                      If (.Stats.ELV > 47) Then
2370                      AumentoHP = .Stats.UserAtributos(eAtributos.Constitucion) \ 4 - AdicionalHPCazador
2380                      End If
                          
2390                      AumentoHIT = 1
                          'AumentoMANA = 3 * .Stats.UserAtributos(eAtributos.Inteligencia)
2400                      AumentoSTA = AumentoSTMago
                          
2410                      If (.Stats.MaxMAN >= 2000) Then
2420                      AumentoMANA = (3 * .Stats.UserAtributos(eAtributos.Inteligencia)) / 2
2430                      Else
2440                      AumentoMANA = 3 * .Stats.UserAtributos(eAtributos.Inteligencia)
2450                      End If
                                    
2460                  Case eClass.Worker
2470                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
2480                      AumentoHP = RandomNumber(9, 12)
2490                      Case 20
2500                      AumentoHP = RandomNumber(8, 12)
2510                      Case 19
2520                      AumentoHP = RandomNumber(7, 12)
2530                      Case 18
2540                      AumentoHP = RandomNumber(6, 11)
2550                      Case Else
2560                      AumentoHP = RandomNumber(6, .Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
2570                      End Select
                      
2580                      AumentoHIT = 1
2590                      AumentoSTA = AumentoSTTrabajador
                      
                   
2600                  Case eClass.Cleric
2610                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
2620                      AumentoHP = RandomNumber(7, 10)
2630                      Case 20
2640                      AumentoHP = RandomNumber(6, 10)
2650                      Case 19
2660                      AumentoHP = RandomNumber(6, 9)
2670                      Case 18
2680                      AumentoHP = RandomNumber(5, 9)
2690                      Case Else
2700                      AumentoHP = RandomNumber(4, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)
2710                      End Select
                      
2720                  If (.Stats.ELV > 47) Then
2730                      AumentoHP = .Stats.UserAtributos(eAtributos.Constitucion) \ 4
2740                      End If
                      
2750                                 If (.Stats.ELV < 48) Then
2760                      AumentoHIT = 2
2770                      Else
2780                      AumentoHIT = 1
2790                      End If
                          
2800                If (.Stats.ELV < 48) Then
2810                AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
2820                Else
2830                AumentoMANA = 1 * .Stats.UserAtributos(eAtributos.Inteligencia)
2840                End If
                      
2850                      AumentoSTA = AumentoSTDef
                      
2860                  Case eClass.Druid
2870                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
2880                      AumentoHP = RandomNumber(7, 10)
2890                      Case 20
2900                      AumentoHP = RandomNumber(6, 10)
2910                      Case 19
2920                      AumentoHP = RandomNumber(6, 9)
2930                      Case 18
2940                      AumentoHP = RandomNumber(5, 9)
2950                      Case Else
2960                      AumentoHP = RandomNumber(4, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)
2970                      End Select
                  
2980                  If (.Stats.ELV > 47) Then
2990                      AumentoHP = .Stats.UserAtributos(eAtributos.Constitucion) \ 4
3000                      End If
                      
3010                                 If (.Stats.ELV < 48) Then
3020                      AumentoHIT = 2
3030                      Else
3040                      AumentoHIT = 1
3050                      End If
                          
3060                If (.Stats.ELV < 48) Then
3070                AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
3080                Else
3090                AumentoMANA = 1 * .Stats.UserAtributos(eAtributos.Inteligencia)
3100                End If
                    
3110                      AumentoSTA = AumentoSTDef
                      
3120                  Case eClass.Assasin
3130                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
3140                      AumentoHP = RandomNumber(7, 10)
3150                      Case 20
3160                      AumentoHP = RandomNumber(6, 10)
3170                      Case 19
3180                      AumentoHP = RandomNumber(6, 9)
3190                      Case 18
3200                      AumentoHP = RandomNumber(5, 9)
3210                      Case Else
3220                      AumentoHP = RandomNumber(4, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)
3230                      End Select
                      
3240                                  If (.Stats.ELV > 47) Then
3250                      AumentoHP = .Stats.UserAtributos(eAtributos.Constitucion) \ 4
3260                      End If
                      
3270                                 If (.Stats.ELV < 48) Then
3280                      AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
3290                      Else
3300                      AumentoHIT = 1
3310                      End If
                          
3320                If (.Stats.ELV < 48) Then
3330                AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
3340                Else
3350                AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
3360                End If
                      
3370                      AumentoSTA = AumentoSTDef
                      
3380                  Case eClass.Bard
3390                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
3400                      AumentoHP = RandomNumber(7, 10)
3410                      Case 20
3420                      AumentoHP = RandomNumber(6, 10)
3430                      Case 19
3440                      AumentoHP = RandomNumber(6, 9)
3450                      Case 18
3460                      AumentoHP = RandomNumber(5, 9)
3470                      Case Else
3480                      AumentoHP = RandomNumber(4, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)
3490                      End Select
                      
3500                  If (.Stats.ELV > 47) Then
3510                      AumentoHP = .Stats.UserAtributos(eAtributos.Constitucion) \ 4
3520                      End If
                      
3530                                 If (.Stats.ELV < 48) Then
3540                      AumentoHIT = 2
3550                      Else
3560                      AumentoHIT = 1
3570                      End If
                          
3580                If (.Stats.ELV < 48) Then
3590                AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
3600                Else
3610                AumentoMANA = 1 * .Stats.UserAtributos(eAtributos.Inteligencia)
3620                End If
                    
3630                      AumentoSTA = AumentoSTDef
                                      
                                      
                                               ' Case eClass.Bandit
                                            '    Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                         ' Case 21
                          'AumentoHP = RandomNumber(9, 11)
                         ' Case 20
                          'AumentoHP = RandomNumber(8, 11)
                          'Case 19
                          'AumentoHP = RandomNumber(7, 11)
                          'Case 18
                          'AumentoHP = RandomNumber(6, 11)
                          'Case Else
                          'AumentoHP = RandomNumber(4, .Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPCazador
                         ' End Select
                          
                                        '  If (.Stats.ELV > 47) Then
                         ' AumentoHP = .Stats.UserAtributos(eAtributos.Constitucion) \ 4 + AdicionalHPCazador
                         ' End If
                      
                           '          If (.Stats.ELV < 48) Then
                         ' AumentoHIT = AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                          'Else
                          'AumentoHIT = 1
                         ' End If
                          
                   ' If (.Stats.ELV < 48) Then
                    'AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia) / 3 * 2
                    'Else
                    'AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia) / 3
                    'End If
                    
                         ' AumentoSTA = AumentoStBandido
                                      
3640                  Case Else
3650                     Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
3660                      AumentoHP = RandomNumber(6, 8)
3670                      Case 20
3680                      AumentoHP = RandomNumber(5, 8)
3690                      Case 19
3700                      AumentoHP = RandomNumber(4, 8)
3710                      Case 18
3720                      AumentoHP = RandomNumber(3, 8)
3730                      Case Else
3740                      AumentoHP = RandomNumber(5, .Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
3750                      End Select
                      
3760                      AumentoHIT = 2
3770                      AumentoSTA = AumentoSTDef
3780              End Select
                  
                  'Actualizamos HitPoints
3790              .Stats.MaxHp = .Stats.MaxHp + AumentoHP
3800              If .Stats.MaxHp > STAT_MAXHP Then .Stats.MaxHp = STAT_MAXHP
                  
                  'Actualizamos Stamina
3810              .Stats.MaxSta = .Stats.MaxSta + AumentoSTA
3820              If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA
                  
                  'Actualizamos Mana
3830              .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA
3840              If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN
                  
                  'Actualizamos Golpe Máximo
3850              .Stats.MaxHIT = .Stats.MaxHIT + AumentoHIT
3860              If .Stats.ELV < 36 Then
3870                  If .Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then _
                          .Stats.MaxHIT = STAT_MAXHIT_UNDER36
3880              Else
3890                  If .Stats.MaxHIT > STAT_MAXHIT_OVER36 Then _
                          .Stats.MaxHIT = STAT_MAXHIT_OVER36
3900              End If
                  
                  'Actualizamos Golpe Mínimo
3910              .Stats.MinHIT = .Stats.MinHIT + AumentoHIT
3920              If .Stats.ELV < 36 Then
3930                  If .Stats.MinHIT > STAT_MAXHIT_UNDER36 Then _
                          .Stats.MinHIT = STAT_MAXHIT_UNDER36
3940              Else
3950                  If .Stats.MinHIT > STAT_MAXHIT_OVER36 Then _
                          .Stats.MinHIT = STAT_MAXHIT_OVER36
3960              End If
                  
                  'Notificamos al user
                  If .Stats.ELV >= 48 Then
3970              If AumentoHP > 0 Then
3980                  Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
3990              End If
4000              If AumentoSTA > 0 Then
4010                  Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoSTA & " puntos de energía.", FontTypeNames.FONTTYPE_INFO)
4020              End If
4030              If AumentoMANA > 0 Then
4040                  Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoMANA & " puntos de maná.", FontTypeNames.FONTTYPE_INFO)
4050              End If
4060              If AumentoHIT > 0 Then
4070                  Call WriteConsoleMsg(UserIndex, "Tu golpe máximo aumentó en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
4080                  Call WriteConsoleMsg(UserIndex, "Tu golpe mínimo aumentó en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
4090              End If
                End If
                  
4100              Call LogDesarrollo(.Name & " paso a nivel " & .Stats.ELV & " gano HP: " & AumentoHP)
                  
4110              .Stats.MinHp = .Stats.MaxHp

                      'If user is in a party, we modify the variable p_sumaniveleselevados
4120                  Call mdParty.ActualizarSumaNivelesElevados(UserIndex)
                          'If user reaches lvl 25 and he is in a guild, we check the guild's alignment and expulses the user if guild has factionary alignment
              
4130              If .Stats.ELV = 25 Then
4140                  GI = .GuildIndex
4150                  If GI > 0 Then
4160                      If modGuilds.GuildAlignment(GI) = "Del Mal" Or modGuilds.GuildAlignment(GI) = "Real" Then
                              'We get here, so guild has factionary alignment, we have to expulse the user
4170                          Call modGuilds.m_EcharMiembroDeClan(-1, .Name)
4180                          Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
4190                          Call WriteConsoleMsg(UserIndex, "¡Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearás! Por esta razón, hasta tanto no te enlistes en la facción bajo la cual tu clan está alineado, estarás excluído del mismo.", FontTypeNames.FONTTYPE_GUILD)
4200                      End If
4210                  End If
4220              End If

4230          Loop
              
              'If it ceased to be a newbie, remove newbie items and get char away from newbie dungeon
4240          If Not EsNewbie(UserIndex) And WasNewbie Then
4250              Call QuitarNewbieObj(UserIndex)
4260              If UCase$(MapInfo(.Pos.map).Restringir) = "NEWBIE" Then
4270                  Call WarpUserChar(UserIndex, 1, 50, 50, True)
4280                  Call WriteConsoleMsg(UserIndex, "Debes abandonar el Dungeon Newbie.", FontTypeNames.FONTTYPE_INFO)
4290              End If
4300          End If
              
              
4360      End With
          
4370      Call WriteUpdateUserStats(UserIndex)
    
    On Error GoTo 0
    Exit Sub

CheckUserLevel_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckUserLevel, line " & Erl & "."

End Sub

Public Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    PuedeAtravesarAgua = UserList(UserIndex).flags.Navegando = 1 _
                    Or UserList(UserIndex).flags.Vuela = 1
End Function

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading)
'*************************************************
'Author: Unknown
'Last modified: 13/07/2009
'Moves the char, sending the message to everyone in range.
'30/03/2009: ZaMa - Now it's legal to move where a casper is, changing its pos to where the moving char was.
'28/05/2009: ZaMa - When you are moved out of an Arena, the resurrection safe is activated.
'13/07/2009: ZaMa - Now all the clients don't know when an invisible admin moves, they force the admin to move.
'13/07/2009: ZaMa - Invisible admins aren't allowed to force dead characater to move
'*************************************************
    Dim nPos As WorldPos
    Dim sailing As Boolean
    Dim CasperIndex As Integer
    Dim CasperHeading As eHeading
    Dim isAdminInvi As Boolean
    
    sailing = PuedeAtravesarAgua(UserIndex)
    nPos = UserList(UserIndex).Pos
    Call HeadtoPos(nHeading, nPos)
        
    isAdminInvi = (UserList(UserIndex).flags.AdminInvisible = 1)
    
    If MoveToLegalPos(UserList(UserIndex).Pos.map, nPos.X, nPos.Y, sailing, Not sailing) Then
        'si no estoy solo en el mapa...
        If MapInfo(UserList(UserIndex).Pos.map).NumUsers > 1 Then
               
            CasperIndex = MapData(UserList(UserIndex).Pos.map, nPos.X, nPos.Y).UserIndex
            'Si hay un usuario, y paso la validacion, entonces es un casper
            If CasperIndex > 0 Then
                ' Los admins invisibles no pueden patear caspers
                If Not isAdminInvi Then
                    
                    If TriggerZonaPelea(UserIndex, CasperIndex) = TRIGGER6_PROHIBE Then
                        If UserList(CasperIndex).flags.SeguroResu = False Then
                            UserList(CasperIndex).flags.SeguroResu = True
                            Call WriteMultiMessage(CasperIndex, eMessages.ResuscitationSafeOn)
                        End If
                    End If
    
                    With UserList(CasperIndex)
                        CasperHeading = InvertHeading(nHeading)
                        Call HeadtoPos(CasperHeading, .Pos)
                    
                        ' Si es un admin invisible, no se avisa a los demas clientes
                        If Not .flags.AdminInvisible = 1 Then _
                            Call SendData(SendTarget.ToPCAreaButIndex, CasperIndex, PrepareMessageCharacterMove(.Char.CharIndex, .Pos.X, .Pos.Y))
                        
                        Call WriteForceCharMove(CasperIndex, CasperHeading)
                            
                        'Update map and char
                        .Char.Heading = CasperHeading
                        MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex = CasperIndex
                    End With
                
                    'Actualizamos las áreas de ser necesario
                    Call ModAreas.CheckUpdateNeededUser(CasperIndex, CasperHeading)
                End If
            End If
            
            ' Si es un admin invisible, no se avisa a los demas clientes
            If Not isAdminInvi Then _
                Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, nPos.X, nPos.Y))
            
        End If
        
        ' Los admins invisibles no pueden patear caspers
        If Not (isAdminInvi And (CasperIndex <> 0)) Then
            Dim oldUserIndex As Integer
            
            With UserList(UserIndex)
                oldUserIndex = MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex
                
                ' Si no hay intercambio de pos con nadie
                If oldUserIndex = UserIndex Then
                    MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex = 0
                End If
                
                .Pos = nPos
                .Char.Heading = nHeading
                MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                
                If HayCura(UserIndex) Then Call Autoresurreccion(UserIndex)
                
                Call DoTileEvents(UserIndex, .Pos.map, .Pos.X, .Pos.Y)
            End With
            
            'Actualizamos las áreas de ser necesario
            Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading)
            
            ' Invocaciones
            If PuedeRealizarInvocacion(UserIndex) Then
                    
                Dim Inv As Byte
                 Inv = Invocation.InvocacionIndex(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
                    
                If Inv > 0 Then
                    If Invocaciones(Inv).Activo = 0 Then
                        Invocation.RealizarInvocacion UserIndex, Inv
                    End If
                End If
            End If

        Else
            Call WritePosUpdate(UserIndex)
        End If
    Else
        Call WritePosUpdate(UserIndex)
    End If
    
    If UserList(UserIndex).Counters.Trabajando Then _
        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

    If UserList(UserIndex).Counters.Ocultando Then _
        UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1

    'Montando
    If UserList(UserIndex).flags.Montando Then SendData SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(RandomNumber(215, 219), UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    
End Sub

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading
'*************************************************
'Author: ZaMa
'Last modified: 30/03/2009
'Returns the heading opposite to the one passed by val.
'*************************************************
    Select Case nHeading
        Case eHeading.EAST
            InvertHeading = WEST
        Case eHeading.WEST
            InvertHeading = EAST
        Case eHeading.SOUTH
            InvertHeading = NORTH
        Case eHeading.NORTH
            InvertHeading = SOUTH
    End Select
End Function

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Object As UserObj)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    UserList(UserIndex).Invent.Object(Slot) = Object
    Call WriteChangeInventorySlot(UserIndex, Slot)
End Sub
Sub Autoresurreccion(ByVal UserIndex As Integer)
'******************************
'Adaptacion a 13.0: Kaneidra
'Last Modification: 15/05/2012
'******************************
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call RevivirUsuario(UserIndex)
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
        Call WriteUpdateMana(UserIndex)
        Call WriteUpdateFollow(UserIndex)
        UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
        Call WriteUpdateHP(UserIndex)
        Call WriteUpdateFollow(UserIndex)
        Call WriteConsoleMsg(UserIndex, "El sacerdote te ha resucitado y curado.", FontTypeNames.FONTTYPE_INFO)
    End If
 
    If UserList(UserIndex).Stats.MinHp < UserList(UserIndex).Stats.MaxHp Then
        UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
        Call WriteUpdateHP(UserIndex)
        Call WriteUpdateFollow(UserIndex)
        Call WriteConsoleMsg(UserIndex, "El sacerdote te ha curado.", FontTypeNames.FONTTYPE_INFO)
    End If
    
 
    If UserList(UserIndex).flags.Envenenado = 1 Then UserList(UserIndex).flags.Envenenado = 0
 
End Sub

Function NextOpenCharIndex() As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim LoopC As Long
    
    For LoopC = 1 To MAXCHARS
        If CharList(LoopC) = 0 Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            
            If LoopC > LastChar Then _
                LastChar = LoopC
            
            Exit Function
        End If
    Next LoopC
End Function

Function NextOpenUser() As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim LoopC As Long
    
    For LoopC = 1 To MaxUsers + 1
        If LoopC > MaxUsers Then Exit For
        If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
    Next LoopC
    
    NextOpenUser = LoopC
End Function

Public Sub SendUserStatsTxt(ByVal SendIndex As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim GuildI As Integer
    
    With UserList(UserIndex)
        Call WriteConsoleMsg(SendIndex, "Estadísticas de: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Nivel: " & .Stats.ELV & "  EXP: " & .Stats.Exp & "/" & .Stats.ELU, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Salud: " & .Stats.MinHp & "/" & .Stats.MaxHp & "  Maná: " & .Stats.MinMAN & "/" & .Stats.MaxMAN & "  Energía: " & .Stats.MinSta & "/" & .Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
        
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call WriteConsoleMsg(SendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHIT & "/" & .Stats.MaxHIT & " (" & ObjData(.Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(.Invent.WeaponEqpObjIndex).MaxHIT & ")", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(SendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHIT & "/" & .Stats.MaxHIT, FontTypeNames.FONTTYPE_INFO)
        End If
        
        If .Invent.ArmourEqpObjIndex > 0 Then
            If .Invent.EscudoEqpObjIndex > 0 Then
                Call WriteConsoleMsg(SendIndex, "(CUERPO) Mín Def/Máx Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef + ObjData(.Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef + ObjData(.Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(SendIndex, "(CUERPO) Mín Def/Máx Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(SendIndex, "(CUERPO) Mín Def/Máx Def: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        
        If .Invent.CascoEqpObjIndex > 0 Then
            Call WriteConsoleMsg(SendIndex, "(CABEZA) Mín Def/Máx Def: " & ObjData(.Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(.Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(SendIndex, "(CABEZA) Mín Def/Máx Def: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        
        GuildI = .GuildIndex
        If GuildI > 0 Then
            Call WriteConsoleMsg(SendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_INFO)
            If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(.Name) Then
                Call WriteConsoleMsg(SendIndex, "Status: Líder", FontTypeNames.FONTTYPE_INFO)
            End If
            'guildpts no tienen objeto
        End If
        
        
#If ConUpTime Then
        Dim TempDate As Date
        Dim TempSecs As Long
        Dim tempStr As String
        TempDate = Now - .LogOnTime
        TempSecs = (.UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
        tempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(SendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Total: " & tempStr, FontTypeNames.FONTTYPE_INFO)
#End If
        
        Call WriteConsoleMsg(SendIndex, "Oro: " & .Stats.Gld & "  Posición: " & .Pos.X & "," & .Pos.Y & " en mapa " & .Pos.map, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Dados: " & .Stats.UserAtributos(eAtributos.Fuerza) & ", " & .Stats.UserAtributos(eAtributos.Agilidad) & ", " & .Stats.UserAtributos(eAtributos.Inteligencia) & ", " & .Stats.UserAtributos(eAtributos.Carisma) & ", " & .Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Retos Ganados: " & .Stats.RetosGanados & "", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Retos Perdidos: " & .Stats.RetosPerdidos & "", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Oro Ganado: " & .Stats.OroGanado & "", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Oro Perdido: " & .Stats.OroPerdido & "", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Torneos Ganados: " & .Stats.TorneosGanados & "", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Sub SendUserMiniStatsTxt(ByVal SendIndex As Integer, ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is online.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
'*************************************************
    With UserList(UserIndex)
        Call WriteConsoleMsg(SendIndex, "Pj: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Ciudadanos matados: " & .Faccion.CiudadanosMatados & " Criminales matados: " & .Faccion.CriminalesMatados & " usuarios matados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "NPCs muertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Clase: " & ListaClases(.clase), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Pena: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)
        
        If .Faccion.ArmadaReal = 1 Then
            Call WriteConsoleMsg(SendIndex, "Ejército real desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Ingresó en nivel: " & .Faccion.NivelIngreso & " con " & .Faccion.MatadosIngreso & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.FuerzasCaos = 1 Then
            Call WriteConsoleMsg(SendIndex, "Legión oscura desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Ingresó en nivel: " & .Faccion.NivelIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.RecibioExpInicialReal = 1 Then
            Call WriteConsoleMsg(SendIndex, "Fue ejército real", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.RecibioExpInicialCaos = 1 Then
            Call WriteConsoleMsg(SendIndex, "Fue legión oscura", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        End If
        
        Call WriteConsoleMsg(SendIndex, "Asesino: " & .Reputacion.AsesinoRep, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Noble: " & .Reputacion.NobleRep, FontTypeNames.FONTTYPE_INFO)
        
        If .GuildIndex > 0 Then
            Call WriteConsoleMsg(SendIndex, "Clan: " & GuildName(.GuildIndex), FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal SendIndex As Integer, ByVal charName As String)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is offline.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
'*************************************************
    Dim CharFile As String
    Dim Ban As String
    Dim BanDetailPath As String
    
    BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile) Then
        Call WriteConsoleMsg(SendIndex, "Pj: " & charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Ciudadanos matados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " usuarios matados: " & GetVar(CharFile, "MUERTES", "UserMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "NPCs muertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Clase: " & ListaClases(GetVar(CharFile, "INIT", "Clase")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Pena: " & GetVar(CharFile, "COUNTERS", "PENA"), FontTypeNames.FONTTYPE_INFO)
        
        If CByte(GetVar(CharFile, "FACCIONES", "EjercitoReal")) = 1 Then
            Call WriteConsoleMsg(SendIndex, "Ejército real desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Ingresó en nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")) & " con " & CInt(GetVar(CharFile, "FACCIONES", "MatadosIngreso")) & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "EjercitoCaos")) = 1 Then
            Call WriteConsoleMsg(SendIndex, "Legión oscura desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Ingresó en nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExReal")) = 1 Then
            Call WriteConsoleMsg(SendIndex, "Fue ejército real", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExCaos")) = 1 Then
            Call WriteConsoleMsg(SendIndex, "Fue legión oscura", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        End If

        
        Call WriteConsoleMsg(SendIndex, "Asesino: " & CLng(GetVar(CharFile, "REP", "Asesino")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Noble: " & CLng(GetVar(CharFile, "REP", "Nobles")), FontTypeNames.FONTTYPE_INFO)
        
        If IsNumeric(GetVar(CharFile, "Guild", "GUILDINDEX")) Then
            Call WriteConsoleMsg(SendIndex, "Clan: " & modGuilds.GuildName(CInt(GetVar(CharFile, "Guild", "GUILDINDEX"))), FontTypeNames.FONTTYPE_INFO)
        End If
        
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call WriteConsoleMsg(SendIndex, "Ban: " & Ban, FontTypeNames.FONTTYPE_INFO)
        
        If Ban = "1" Then
            Call WriteConsoleMsg(SendIndex, "Ban por: " & GetVar(CharFile, charName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, charName, "Reason"), FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(SendIndex, "El pj no existe: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub SendUserInvTxt(ByVal SendIndex As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    Dim j As Long
    
    With UserList(UserIndex)
        Call WriteConsoleMsg(SendIndex, .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Tiene " & .Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To .CurrentInventorySlots
            If .Invent.Object(j).objindex > 0 Then
                Call WriteConsoleMsg(SendIndex, "Objeto " & j & " " & ObjData(.Invent.Object(j).objindex).Name & " Cantidad:" & .Invent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)
            End If
        Next j
    End With
End Sub

Sub SendUserInvTxtFromChar(ByVal SendIndex As Integer, ByVal charName As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    Dim j As Long
    Dim CharFile As String, tmp As String
    Dim ObjInd As Long, ObjCant As Long
    
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(SendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, tmp, Asc("-"))
            ObjCant = ReadField(2, tmp, Asc("-"))
            If ObjInd > 0 Then
                Call WriteConsoleMsg(SendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)
            End If
        Next j
    Else
        Call WriteConsoleMsg(SendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub SendUserSkillsTxt(ByVal SendIndex As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next
    Dim j As Integer
    
    Call WriteConsoleMsg(SendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
    
    For j = 1 To NUMSKILLS
        Call WriteConsoleMsg(SendIndex, SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
    Next j
    
    Call WriteConsoleMsg(SendIndex, "SkillLibres:" & UserList(UserIndex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)
End Sub

Private Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not criminal(Npclist(NpcIndex).MaestroUser)
        If EsMascotaCiudadano Then
            Call WriteConsoleMsg(Npclist(NpcIndex).MaestroUser, "¡¡" & UserList(UserIndex).Name & " esta atacando tu mascota!!", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
End Function
Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
'**********************************************
'Author: Unknown
'Last Modification: 02/04/2010
'24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
'24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
'06/28/2008 -> NicoNZ: Los elementales al atacarlos por su amo no se paran más al lado de él sin hacer nada.
'02/04/2010: ZaMa: Un ciuda no se vuelve mas criminal al atacar un npc no hostil.
'**********************************************
    Dim EraCriminal As Boolean
    
    'Guardamos el usuario que ataco el npc.
    Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).Name
    
    'Npc que estabas atacando.
    Dim LastNpcHit As Integer
    LastNpcHit = UserList(UserIndex).flags.NPCAtacado
    'Guarda el NPC que estas atacando ahora.
    UserList(UserIndex).flags.NPCAtacado = NpcIndex
    
    'Revisamos robo de npc.
    'Guarda el primer nick que lo ataca.
    If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
            End If
        End If
        Npclist(NpcIndex).flags.AttackedFirstBy = UserList(UserIndex).Name
    ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(UserIndex).Name Then
        'Estas robando NPC
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
            End If
        End If
    End If
    
    If Npclist(NpcIndex).MaestroUser > 0 Then
        If Npclist(NpcIndex).MaestroUser <> UserIndex Then
            Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)
        End If
    End If
    
    If EsMascotaCiudadano(NpcIndex, UserIndex) Then
        Call VolverCriminal(UserIndex)
        Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
        Npclist(NpcIndex).Hostile = 1
    Else
        EraCriminal = criminal(UserIndex)
        
        'Reputacion
        If Npclist(NpcIndex).Stats.Alineacion = 0 Then
           If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                Call VolverCriminal(UserIndex)
           End If
        
        ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
           UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlCAZADOR / 2
           If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
            UserList(UserIndex).Reputacion.PlebeRep = MAXREP
        End If
        
        If Npclist(NpcIndex).MaestroUser <> UserIndex Then
            'hacemos que el npc se defienda
            Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
            Npclist(NpcIndex).Hostile = 1
        End If
        
        If EraCriminal And Not criminal(UserIndex) Then
            Call VolverCiudadano(UserIndex)
        End If
    End If
End Sub
Public Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1 Then
            PuedeApuñalar = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR _
                        Or UserList(UserIndex).clase = eClass.Assasin
        End If
    End If
End Function

Public Function PuedeAcuchillar(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 25/01/2010 (ZaMa)
'
'***************************************************
    
    With UserList(UserIndex)
        If .clase = eClass.Pirat Then
            If .Invent.WeaponEqpObjIndex > 0 Then
                PuedeAcuchillar = (ObjData(.Invent.WeaponEqpObjIndex).Acuchilla = 1)
            End If
        End If
    End With
    
End Function

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer, ByVal Acerto As Boolean)
'*************************************************
'Author: Unknown
'Last modified: 30/01/2012
'11/19/2009 Pato   - Implement the new system to train the skills.
'30/01/2012 maTih - Modifico la subida de skills fáciles.
'*************************************************
    With UserList(UserIndex)
        If .flags.Hambre = 0 And .flags.Sed = 0 Then
            If .Counters.AsignedSkills < 10 Then
                If Not .flags.UltimoMensaje = 7 Then
                    Call WriteConsoleMsg(UserIndex, "Para poder entrenar un skill debes asignar los 10 skills iniciales.", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 7
                End If
                
                Exit Sub
            End If
                
            With .Stats
                If .UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
                
                Dim Lvl As Integer
                Lvl = .ELV
                
                If Lvl > UBound(LevelSkill) Then Lvl = UBound(LevelSkill)
                
                If .UserSkills(Skill) >= LevelSkill(Lvl).LevelValue Then Exit Sub
                
                .ExpSkills(Skill) = .EluSkills(Skill)
                
                If .ExpSkills(Skill) >= .EluSkills(Skill) Then
                    .UserSkills(Skill) = .UserSkills(Skill) + 1
                    Call WriteConsoleMsg(UserIndex, "¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto! Ahora tienes " & .UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
                    
                    .Exp = .Exp + 50
                    If .Exp > MAXEXP Then .Exp = MAXEXP
                    
                    Call WriteConsoleMsg(UserIndex, "¡Has ganado 50 puntos de experiencia!", FontTypeNames.FONTTYPE_FIGHT)
                    
                    Call WriteUpdateExp(UserIndex)
                    Call CheckUserLevel(UserIndex)
                    Call CheckEluSkill(UserIndex, Skill, False)
                End If
            End With
        End If
    End With
End Sub


''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'
Sub UserDie(ByVal UserIndex As Integer, Optional ByVal AttackerIndex As Integer = 0)
        '************************************************
        'Author: Uknown
        'Last Modified: 12/01/2010 (ZaMa)
        '04/15/2008: NicoNZ - Ahora se resetea el counter del invi
        '13/02/2009: ZaMa - Ahora se borran las mascotas cuando moris en agua.
        '27/05/2009: ZaMa - El seguro de resu no se activa si estas en una arena.
        '21/07/2009: Marco - Al morir se desactiva el comercio seguro.
        '16/11/2009: ZaMa - Al morir perdes la criatura que te pertenecia.
        '27/11/2009: Budi - Al morir envia los atributos originales.
        '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando mueren.
        '************************************************
       
        '<EhHeader>
        On Error GoTo UserDie_Err
 
        '</EhHeader>
        Dim i  As Long
        Dim aN As Integer
 
100     With UserList(UserIndex)
 
            'Sonido
102         If .Genero = eGenero.Mujer Then
104             Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MUERTE_MUJER)
            Else
106             Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MUERTE_HOMBRE)
 
            End If
     
            'Quitar el dialogo del user muerto
108         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
     
110         .Stats.MinHp = 0
112         .Stats.MinSta = 0
114         .flags.AtacadoPorUser = 0
116         .flags.Envenenado = 0
118         .flags.Muerto = 1
     
120         .Counters.Trabajando = 0
     
            ' No se activa en arenas
122         If TriggerZonaPelea(UserIndex, UserIndex) <> TRIGGER6_PERMITE Then
124             .flags.SeguroResu = True
126             Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn) 'Call WriteResuscitationSafeOn(UserIndex)
            Else
128             .flags.SeguroResu = False
130             Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff) 'Call WriteResuscitationSafeOff(UserIndex)
 
            End If
     
132         aN = .flags.AtacadoPorNpc
 
134         If aN > 0 Then
136             Npclist(aN).Movement = Npclist(aN).flags.OldMovement
138             Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
140             Npclist(aN).flags.AttackedBy = vbNullString
 
            End If
     
142         aN = .flags.NPCAtacado
 
144         If aN > 0 Then
146             If Npclist(aN).flags.AttackedFirstBy = .Name Then
148                 Npclist(aN).flags.AttackedFirstBy = vbNullString
 
                End If
 
            End If
 
150         .flags.AtacadoPorNpc = 0
152         .flags.NPCAtacado = 0
     
154         Call PerdioNpc(UserIndex, False)
     
     
        If .flags.Montando = 1 Then
            .flags.Montando = 0
            Call WriteMontateToggle(UserIndex)
            End If
     
            '<<<< Atacable >>>>
156         If .flags.AtacablePor > 0 Then
158             .flags.AtacablePor = 0
160             Call RefreshCharStatus(UserIndex)
 
            End If
     
            '<<<< Paralisis >>>>
162         If .flags.Paralizado = 1 Then
164             .flags.Paralizado = 0
166             Call WriteParalizeOK(UserIndex)
 
            End If
     
            '<<< Estupidez >>>
168         If .flags.Estupidez = 1 Then
170             .flags.Estupidez = 0
172             Call WriteDumbNoMore(UserIndex)
 
            End If
     
            '<<<< Descansando >>>>
174         If .flags.Descansar Then
176             .flags.Descansar = False
178             Call WriteRestOK(UserIndex)
 
            End If
     
            '<<<< Meditando >>>>
180         If .flags.Meditando Then
182             .flags.Meditando = False
184             Call WriteMeditateToggle(UserIndex)
 
            End If
     
            '<<<< Invisible >>>>
186         If .flags.invisible = 1 Or .flags.Oculto = 1 Then
188             .flags.Oculto = 0
190             .flags.invisible = 0
192             .Counters.TiempoOculto = 0
194             .Counters.Invisibilidad = 0
         
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
196             Call SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
 
            End If
     
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
198         Call SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
200         If TriggerZonaPelea(UserIndex, UserIndex) <> eTrigger6.TRIGGER6_PERMITE Then
 
                ' << Si es newbie no pierde el inventario >>
202             If Not EsNewbie(UserIndex) Then
204                 If MapInfo(.Pos.map).SeCaenItems = False Then
206                     Call TirarTodo(UserIndex)
 
                    End If
 
                Else
 
208                 If EsNewbie(UserIndex) And MapInfo(.Pos.map).SeCaenItems = False Then
210                     Call TirarTodosLosItemsNoNewbies(UserIndex)
 
                    End If
 
                End If
 
            End If
            ' DESEQUIPA TODOS LOS OBJETOS
            'desequipar armadura
212         If .Invent.ArmourEqpObjIndex > 0 Then
214             Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
 
            End If
     
            'desequipar arma
216         If .Invent.WeaponEqpObjIndex > 0 Then
218             Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
 
            End If
     
            'desequipar casco
220         If .Invent.CascoEqpObjIndex > 0 Then
222             Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
 
            End If
     
            'desequipar herramienta
224         If .Invent.AnilloEqpSlot > 0 Then
226             Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)
 
            End If
     
            'desequipar municiones
228         If .Invent.MunicionEqpObjIndex > 0 Then
230             Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
 
            End If
     
            'desequipar escudo
232         If .Invent.EscudoEqpObjIndex > 0 Then
234             Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
 
            End If
     
            ' << Reseteamos los posibles FX sobre el personaje >>
236         If .Char.loops = INFINITE_LOOPS Then
238             .Char.FX = 0
240             .Char.loops = 0
 
            End If
     
            ' << Restauramos el mimetismo
242         If .flags.Mimetizado = 1 Then
244             .Char.body = .CharMimetizado.body
246             .Char.Head = .CharMimetizado.Head
248             .Char.CascoAnim = .CharMimetizado.CascoAnim
250             .Char.ShieldAnim = .CharMimetizado.ShieldAnim
252             .Char.WeaponAnim = .CharMimetizado.WeaponAnim
254             .Counters.Mimetismo = 0
256             .flags.Mimetizado = 0
                ' Puede ser atacado por npcs (cuando resucite)
258             .flags.Ignorado = False
 
            End If
     
            ' << Restauramos los atributos >>
260         If .flags.TomoPocion = True Then
 
262             For i = 1 To 5
264                 .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
266             Next i
 
            End If
     
            '<< Cambiamos la apariencia del char >>
268         If .flags.Navegando = 0 Then
270             If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
272                 .Char.body = iCuerpoMuertoCrimi
274                 .Char.Head = iCabezaMuertoCrimi
276                 .Char.ShieldAnim = NingunEscudo
278                 .Char.WeaponAnim = NingunArma
280                 .Char.CascoAnim = NingunCasco
                Else
282                 .Char.body = iCuerpoMuerto
284                 .Char.Head = iCabezaMuerto
286                 .Char.ShieldAnim = NingunEscudo
288                 .Char.WeaponAnim = NingunArma
290                 .Char.CascoAnim = NingunCasco
 
                End If
 
            Else
292             .Char.body = iFragataFantasmal
 
            End If
     
294         For i = 1 To MAXMASCOTAS
 
296             If .MascotasIndex(i) > 0 Then
298                 Call MuereNpc(.MascotasIndex(i), 0)
                    ' Si estan en agua o zona segura
                Else
300                 .MascotasType(i) = 0
 
                End If
 
302         Next i
     
304         .NroMascotas = 0
     
    
      
            '<< Actualizamos clientes >>
306         Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
308         Call WriteUpdateUserStats(UserIndex)
310         Call WriteUpdateStrenghtAndDexterity(UserIndex)
 
            '<<Castigos por party>>
312         If .PartyIndex > 0 Then
314             Call mdParty.ObtenerExito(UserIndex, .Stats.ELV * -10 * mdParty.CantMiembros(UserIndex), .Pos.map, _
                        .Pos.X, .Pos.Y)
 
            End If
     
            '<<Cerramos comercio seguro>>
316         Call LimpiarComercioSeguro(UserIndex)
     
318         If UserList(UserIndex).flags.automatico = True Then
320             Call Rondas_UsuarioMuere(UserIndex)
 
            End If
     
322         If UserList(UserIndex).flags.Plantico = True Then
324             Call Rondas_PlantadorMuere(UserIndex)
 
            End If
     
           Call eventDie(UserIndex)
     
     
     
                     'Estaba en los JDH?
        If .hungry Then
           Call Mod_Jdh.MuereUser(UserIndex)
        End If
        
            'Estaba en death?
326        ' If .death = True Then
328            ' Call ModDeath.MuereUser(UserIndex)
 
            'End If
     
            'Evento 1vs1
330         If .Reto1vs1.RetoIndex <> 0 Then Retos1vs1.Muere UserIndex
     
            ' Hay que teletransportar?
            Dim Mapa As Integer
334         Mapa = .Pos.map
            Dim MapaTelep As Integer
336         MapaTelep = MapInfo(Mapa).OnDeathGoTo.map
     
338         If MapaTelep <> 0 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡¡Tu estado no te permite permanecer en el mapa!!!", FontTypeNames.FONTTYPE_INFOBOLD)
340             Call WarpUserChar(UserIndex, MapaTelep, MapInfo(Mapa).OnDeathGoTo.X, MapInfo(Mapa).OnDeathGoTo.Y, _
                        True, True)
 
            End If
            
            ' ¿Damos el poder?
            If StrComp(GreatPower.CurrentUser, UCase$(.Name)) = 0 Then
                Check_GreatPower UserIndex, AttackerIndex
            End If
            
            If .flags.SlotEvent > 0 Then
                If Events(.flags.SlotEvent).Modality = CastleMode Then
                    .Counters.TimeCastleMode = 10
                End If
                
                If Events(.flags.SlotEvent).Modality = DeathMatch Then
                    Call DeathMatch_UserDie(.flags.SlotEvent, UserIndex)
                End If

            End If
            
            Dim SlotEvent As Byte
            SlotEvent = .flags.SlotEvent
                
            If SlotEvent > 0 Then
                Select Case Events(SlotEvent).Modality
                    Case eModalityEvent.CastleMode
                            .Counters.TimeCastleMode = 3
                    Case eModalityEvent.DeathMatch
                            Call DeathMatch_UserDie(SlotEvent, UserIndex)
                    Case eModalityEvent.Enfrentamientos
                            Fight_UserDie SlotEvent, .flags.SlotUserEvent, AttackerIndex
                End Select
            End If
            
            
        End With
 
        '<EhFooter>
        Exit Sub
 
UserDie_Err:
        LogError Err.Description & vbCrLf & "UserDie " & "at line " & Erl
 
        '</EhFooter>
End Sub
Public Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 13/07/2010
'13/07/2010: ZaMa - Los matados en estado atacable ya no suman frag.
'***************************************************
 
    If EsNewbie(Muerto) Then Exit Sub
       
    With UserList(Atacante)
        If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
       
        If criminal(Muerto) Then
            If .flags.LastCrimMatado <> UserList(Muerto).Name Then
                .flags.LastCrimMatado = UserList(Muerto).Name
                If .Faccion.CriminalesMatados < MAXUSERMATADOS Then _
                    .Faccion.CriminalesMatados = .Faccion.CriminalesMatados + 1
            End If
           
            If .Faccion.RecibioExpInicialCaos = 1 And UserList(Muerto).Faccion.FuerzasCaos = 1 Then
                .Faccion.Reenlistadas = 200  'jaja que trucho
               
                'con esto evitamos que se vuelva a reenlistar
            End If
        Else
            If .flags.LastCiudMatado <> UserList(Muerto).Name Then
                .flags.LastCiudMatado = UserList(Muerto).Name
                If .Faccion.CiudadanosMatados < MAXUSERMATADOS Then _
                    .Faccion.CiudadanosMatados = .Faccion.CiudadanosMatados + 1
            End If
        End If
       
        If .Stats.UsuariosMatados < MAXUSERMATADOS Then _
            .Stats.UsuariosMatados = .Stats.UsuariosMatados + 1
    End With
End Sub

Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef Obj As Obj, _
              ByRef PuedeAgua As Boolean, ByRef PuedeTierra As Boolean)
'**************************************************************
'Author: Unknown
'Last Modify Date: 18/09/2010
'23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
'18/09/2010: ZaMa - Aplico optimizacion de busqueda de tile libre en forma de rombo.
'**************************************************************
On Error GoTo Errhandler

    Dim Found As Boolean
    Dim LoopC As Integer
    Dim tX As Long
    Dim tY As Long
    
    nPos = Pos
    tX = Pos.X
    tY = Pos.Y
    
    LoopC = 1
    
    ' La primera posicion es valida?
    If LegalPos(Pos.map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra, True) Then
        
        If Not HayObjeto(Pos.map, nPos.X, nPos.Y, Obj.objindex, Obj.Amount) Then
            Found = True
        End If
        
    End If
    
    ' Busca en las demas posiciones, en forma de "rombo"
    If Not Found Then
        While (Not Found) And LoopC <= 16
            If RhombLegalTilePos(Pos, tX, tY, LoopC, Obj.objindex, Obj.Amount, PuedeAgua, PuedeTierra) Then
                nPos.X = tX
                nPos.Y = tY
                Found = True
            End If
        
            LoopC = LoopC + 1
        Wend
        
    End If
    
    If Not Found Then
        nPos.X = 0
        nPos.Y = 0
    End If
    
    Exit Sub
    
Errhandler:
    Call LogError("Error en Tilelibre. Error: " & Err.Number & " - " & Err.Description)
End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, _
ByVal FX As Boolean, Optional ByVal Teletransported As Boolean)
'**************************************************************
'Author: Unknown
'Last Modify Date: 11/23/2010
'15/07/2009 - ZaMa: Automatic toogle navigate after warping to water.
'13/11/2009 - ZaMa: Now it's activated the timer which determines if the npc can atacak the user.
'16/09/2010 - ZaMa: No se pierde la visibilidad al cambiar de mapa al estar navegando invisible.
'11/23/2010 - C4b3z0n: Ahora si no se permite Invi o Ocultar en el mapa al que cambias, te lo saca
'**************************************************************
    Dim OldMap As Integer
    Dim OldX As Integer
    Dim OldY As Integer
    
    With UserList(UserIndex)
        'Quitar el dialogo
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
        OldMap = .Pos.map
        OldX = .Pos.X
        OldY = .Pos.Y

        Call EraseUserChar(UserIndex, .flags.AdminInvisible = 1)
        
        If OldMap <> map Then
            Call WriteChangeMap(UserIndex, map, MapInfo(.Pos.map).MapVersion)
            
            If .flags.Privilegios And PlayerType.User Then 'El chequeo de invi/ocultar solo afecta a Usuarios (C4b3z0n)
                Dim AhoraVisible As Boolean 'Para enviar el mensaje de invi y hacer visible (C4b3z0n)
                Dim WasInvi As Boolean
                'Chequeo de flags de mapa por invisibilidad (C4b3z0n)
                If MapInfo(map).InviSinEfecto > 0 And .flags.invisible = 1 Then
                    .flags.invisible = 0
                    .Counters.Invisibilidad = 0
                    AhoraVisible = True
                    WasInvi = True 'si era invi, para el string
                End If
                'Chequeo de flags de mapa por ocultar (C4b3z0n)
                If MapInfo(map).OcultarSinEfecto > 0 And .flags.Oculto = 1 Then
                    AhoraVisible = True
                    .flags.Oculto = 0
                    .Counters.TiempoOculto = 0
                End If
                
                If AhoraVisible Then 'Si no era visible y ahora es, le avisa. (C4b3z0n)
                    Call SetInvisible(UserIndex, .Char.CharIndex, False)
                    If WasInvi Then 'era invi
                        Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible ya que no esta permitida la invisibilidad en este mapa.", FontTypeNames.FONTTYPE_INFO)
                    Else 'estaba oculto
                        Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible ya que no esta permitido ocultarse en este mapa.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
            
            Call WritePlayMidi(UserIndex, val(ReadField(1, MapInfo(map).Music, 45)))
            
    
            'Update new Map Users
            MapInfo(map).NumUsers = MapInfo(map).NumUsers + 1
            
            'Update old Map Users
            MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
            If MapInfo(OldMap).NumUsers < 0 Then
                MapInfo(OldMap).NumUsers = 0
            End If
        
            'Si el mapa al que entro NO ES superficial AND en el que estaba TAMPOCO ES superficial, ENTONCES
            Dim nextMap, previousMap As Boolean
            nextMap = IIf(distanceToCities(map).distanceToCity(.Hogar) >= 0, True, False)
            previousMap = IIf(distanceToCities(.Pos.map).distanceToCity(.Hogar) >= 0, True, False)

            If previousMap And nextMap Then '138 => 139 (Ambos superficiales, no tiene que pasar nada)
                'NO PASA NADA PORQUE NO ENTRO A UN DUNGEON.
            ElseIf previousMap And Not nextMap Then '139 => 140 (139 es superficial, 140 no. Por lo tanto 139 es el ultimo mapa superficial)
                .flags.lastMap = .Pos.map
            ElseIf Not previousMap And nextMap Then '140 => 139 (140 es no es superficial, 139 si. Por lo tanto, el último mapa es 0 ya que no esta en un dungeon)
                .flags.lastMap = 0
            ElseIf Not previousMap And Not nextMap Then '140 => 141 (Ninguno es superficial, el ultimo mapa es el mismo de antes)
                .flags.lastMap = .flags.lastMap
            End If
            
            Call WriteRemoveAllDialogs(UserIndex)
        End If
        
        .Pos.X = X
        .Pos.Y = Y
        .Pos.map = map
        
        ' Chequeamos el gran poder
        If StrComp(GreatPower.CurrentUser, UCase$(.Name)) = 0 Then
            mGranPoder.Check_GreatPower (UserIndex)
        End If
        
        Call MakeUserChar(True, map, UserIndex, map, X, Y)
        Call WriteUserCharIndexInServer(UserIndex)
        
        Call DoTileEvents(UserIndex, map, X, Y)
        
        'Force a flush, so user index is in there before it's destroyed for teleporting
        Call FlushBuffer(UserIndex)
        
        'Seguis invisible al pasar de mapa
        If (.flags.invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then
            
            ' No si estas navegando
            If .flags.Navegando = 0 Then
                Call SetInvisible(UserIndex, .Char.CharIndex, True)
            End If
        End If
        
        If Teletransported Then
            If .flags.Traveling = 1 Then
                .flags.Traveling = 0
                .Counters.goHome = 0
                Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
            End If
        End If
        
        If FX And .flags.AdminInvisible = 0 Then 'FX
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
        End If
        
        If .NroMascotas Then Call WarpMascotas(UserIndex)

        ' No puede ser atacado cuando cambia de mapa, por cierto tiempo
        Call IntervaloPermiteSerAtacado(UserIndex, True)
        
        ' Perdes el npc al cambiar de mapa
        Call PerdioNpc(UserIndex, False)
        
        ' Automatic toogle navigate
        If (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero)) = 0 Then
            If HayAgua(.Pos.map, .Pos.X, .Pos.Y) Then
                If .flags.Navegando = 0 Then
                    .flags.Navegando = 1
                        
                    'Tell the client that we are navigating.
                    Call WriteNavigateToggle(UserIndex)
                End If
            Else
                If .flags.Navegando = 1 Then
                    .flags.Navegando = 0
                            
                    'Tell the client that we are navigating.
                    Call WriteNavigateToggle(UserIndex)
                End If
            End If
        End If
      
    End With
End Sub

Public Sub WarpMascotas(ByVal UserIndex As Integer, Optional CanWarp As Boolean)
'************************************************
'Author: Uknown
'Last Modified: 11/05/2009
'13/02/2009: ZaMa - Arreglado respawn de mascotas al cambiar de mapa.
'13/02/2009: ZaMa - Las mascotas no regeneran su vida al cambiar de mapa (Solo entre mapas inseguros).
'11/05/2009: ZaMa - Chequeo si la mascota pueden spwnear para asiganrle los stats.
'************************************************
    Dim i As Integer
    Dim petType As Integer
    Dim PetRespawn As Boolean
    Dim PetTiempoDeVida As Integer
    Dim NroPets As Integer
    Dim InvocadosMatados As Integer
 '   Dim CanWarp As Boolean
    Dim index As Integer
    Dim iMinHP As Integer
    
    NroPets = UserList(UserIndex).NroMascotas
   ' CanWarp = (MapInfo(UserList(UserIndex).Pos.Map).Pk = True)
    
    For i = 1 To MAXMASCOTAS
        index = UserList(UserIndex).MascotasIndex(i)
        
        If index > 0 Then
            ' si la mascota tiene tiempo de vida > 0 significa q fue invocada => we kill it
            If Npclist(index).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(index)
                UserList(UserIndex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
                
                petType = 0
                UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1
            Else
                'Store data and remove NPC to recreate it after warp
                'PetRespawn = Npclist(index).flags.Respawn = 0
                petType = UserList(UserIndex).MascotasType(i)
                'PetTiempoDeVida = Npclist(index).Contadores.TiempoExistencia
                
                ' Guardamos el hp, para restaurarlo uando se cree el npc
                iMinHP = Npclist(index).Stats.MinHp
                
                Call QuitarNPC(index)
                
                ' Restauramos el valor de la variable
                UserList(UserIndex).MascotasType(i) = petType

            End If
        ElseIf UserList(UserIndex).MascotasType(i) > 0 Then
            'Store data and remove NPC to recreate it after warp
            PetRespawn = True
            petType = UserList(UserIndex).MascotasType(i)
            PetTiempoDeVida = 0
        Else
            petType = 0
        End If
        
        If petType > 0 And CanWarp Then
            index = SpawnNpc(petType, UserList(UserIndex).Pos, True, PetRespawn)
            
            'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
            ' Exception: Pets don't spawn in water if they can't swim
            If index = 0 Then
                Call WriteConsoleMsg(UserIndex, "Tus mascotas no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
            Else
                UserList(UserIndex).MascotasIndex(i) = index

                ' Nos aseguramos de que conserve el hp, si estaba dañado
                Npclist(index).Stats.MinHp = IIf(iMinHP = 0, Npclist(index).Stats.MinHp, iMinHP)
            
                Npclist(index).MaestroUser = UserIndex
                Npclist(index).Movement = TipoAI.SigueAmo
                Npclist(index).Target = 0
                Npclist(index).TargetNPC = 0
                Npclist(index).Contadores.TiempoExistencia = PetTiempoDeVida
                Call FollowAmo(index)
            End If
        End If
    Next i
    
    If InvocadosMatados > 0 Then
        Call WriteConsoleMsg(UserIndex, "Pierdes el control de tus mascotas invocadas.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    If Not CanWarp Then
     '   Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    UserList(UserIndex).NroMascotas = NroPets
End Sub

Public Sub WarpMascota(ByVal UserIndex As Integer, ByVal PetIndex As Integer)
'************************************************
'Author: ZaMa
'Last Modified: 18/11/2009
'Warps a pet without changing its stats
'************************************************
    Dim petType As Integer
    Dim NpcIndex As Integer
    Dim iMinHP As Integer
    Dim TargetPos As WorldPos
    
       With UserList(UserIndex)
        
        TargetPos.map = .flags.TargetMap
        TargetPos.X = .flags.TargetX
        TargetPos.Y = .flags.TargetY
        
        NpcIndex = .MascotasIndex(PetIndex)
            
        'Store data and remove NPC to recreate it after warp
        petType = .MascotasType(PetIndex)
        
        ' Guardamos el hp, para restaurarlo cuando se cree el npc
        iMinHP = Npclist(NpcIndex).Stats.MinHp
        
        Call QuitarNPC(NpcIndex)
        
        ' Restauramos el valor de la variable
        .MascotasType(PetIndex) = petType
        .NroMascotas = .NroMascotas + 1
        NpcIndex = SpawnNpc(petType, TargetPos, False, False)
        
        'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
        ' Exception: Pets don't spawn in water if they can't swim
        If NpcIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, "Tu mascota no pueden transitar este sector del mapa, intenta invocarla en otra parte.", FontTypeNames.FONTTYPE_INFO)
        Else
            .MascotasIndex(PetIndex) = NpcIndex

            With Npclist(NpcIndex)
                ' Nos aseguramos de que conserve el hp, si estaba dañado
                .Stats.MinHp = IIf(iMinHP = 0, .Stats.MinHp, iMinHP)
            
                .MaestroUser = UserIndex
                .Movement = TipoAI.SigueAmo
                .Target = 0
                .TargetNPC = 0
            End With
            
            Call FollowAmo(NpcIndex)
        End If
    End With
End Sub


''
' Se inicia la salida de un usuario.
'
' @param    UserIndex   El index del usuario que va a salir

Sub Cerrar_Usuario(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/09/2010
'16/09/2010 - ZaMa: Cuando se va el invi estando navegando, no se saca el invi (ya esta visible).
'***************************************************
    Dim isNotVisible As Boolean
    Dim HiddenPirat As Boolean
    
    With UserList(UserIndex)
        If .flags.UserLogged And Not .Counters.Saliendo Then
            .Counters.Saliendo = True
            .Counters.Salir = IIf((.flags.Privilegios And PlayerType.User) And MapInfo(.Pos.map).Pk, IntervaloCerrarConexion, 0)
            
            isNotVisible = (.flags.Oculto Or .flags.invisible)
            If isNotVisible Then
                .flags.invisible = 0
                .Counters.Invisibilidad = 0
                
                If .flags.Oculto Then
                    If .flags.Navegando = 1 Then
                        If .clase = eClass.Pirat Then
                            ' Pierde la apariencia de fragata fantasmal
                            Call ToggleBoatBody(UserIndex)
                            Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, _
                                                NingunEscudo, NingunCasco)
                            HiddenPirat = True
                        End If
                    End If
                End If
                
                .flags.Oculto = 0
                
                
                ' Para no repetir mensajes
                If Not HiddenPirat Then Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                
                ' Si esta navegando ya esta visible
                If .flags.Navegando = 0 Then
                    Call SetInvisible(UserIndex, .Char.CharIndex, False)
                End If
            End If
            
            If .flags.Traveling = 1 Then
                Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
                .flags.Traveling = 0
                .Counters.goHome = 0
            End If
            
            
            'Call WriteConsoleMsg(UserIndex, "Gracias por jugar Desterium AO.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Cerrando...Se cerrará el juego en " & .Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub


''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/02/08
'
'***************************************************
    If UserList(UserIndex).Counters.Saliendo Then
        ' Is the user still connected?
        If UserList(UserIndex).ConnIDValida Then
            UserList(UserIndex).Counters.Saliendo = False
            UserList(UserIndex).Counters.Salir = 0
            Call WriteConsoleMsg(UserIndex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
        Else
            'Simply reset
            UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(UserIndex).Pos.map).Pk, IntervaloCerrarConexion, 0)
        End If
    End If
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal UserIndex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim ViejoNick As String
    Dim ViejoCharBackup As String
    
    If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
    ViejoNick = UserList(UserIndexDestino).Name
    
    If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
        'hace un backup del char
        ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
        Name CharPath & ViejoNick & ".chr" As ViejoCharBackup
    End If
End Sub

Sub SendUserStatsTxtOFF(ByVal SendIndex As Integer, ByVal Nombre As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If FileExist(CharPath & Nombre & ".chr", vbArchive) = False Then
        Call WriteConsoleMsg(SendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(SendIndex, "Estadísticas de: " & Nombre, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Nivel: " & GetVar(CharPath & Nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & Nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "elu"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Energía: " & GetVar(CharPath & Nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "maxSta"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Salud: " & GetVar(CharPath & Nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxHP") & "  Maná: " & GetVar(CharPath & Nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxMAN"), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(SendIndex, "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "stats", "MaxHIT"), FontTypeNames.FONTTYPE_INFO)

        Call WriteConsoleMsg(SendIndex, "Oro: " & GetVar(CharPath & Nombre & ".chr", "stats", "GLD"), FontTypeNames.FONTTYPE_INFO)
        
#If ConUpTime Then
        Dim TempSecs As Long
        Dim tempStr As String
        TempSecs = GetVar(CharPath & Nombre & ".chr", "INIT", "UpTime")
        tempStr = (TempSecs \ 86400) & " Días, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(SendIndex, "Tiempo Logeado: " & tempStr, FontTypeNames.FONTTYPE_INFO)
#End If
    
    End If
End Sub

Sub SendUserOROTxtFromChar(ByVal SendIndex As Integer, ByVal charName As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim CharFile As String
    
On Error Resume Next
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(SendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(SendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 21/02/2010
'Nacho: Actualiza el tag al cliente
'21/02/2010: ZaMa - Ahora deja de ser atacable si se hace criminal.
'**************************************************************
    With UserList(UserIndex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
            .Reputacion.BurguesRep = 0
            .Reputacion.NobleRep = 0
            .Reputacion.PlebeRep = 0
            .Reputacion.BandidoRep = .Reputacion.BandidoRep + vlASALTO
            If .Reputacion.BandidoRep > MAXREP Then .Reputacion.BandidoRep = MAXREP
            If .Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)
            
            If .flags.AtacablePor > 0 Then .flags.AtacablePor = 0

        End If
    End With
    
    Call RefreshCharStatus(UserIndex)
End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 21/06/2006
'Nacho: Actualiza el tag al cliente.
'**************************************************************
    With UserList(UserIndex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        .Reputacion.LadronesRep = 0
        .Reputacion.BandidoRep = 0
        .Reputacion.AsesinoRep = 0
        .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlASALTO
        If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
    End With
    
    Call RefreshCharStatus(UserIndex)
End Sub

''
'Checks if a given body index is a boat or not.
'
'@param body    The body index to bechecked.
'@return    True if the body is a boat, false otherwise.

Public Function BodyIsBoat(ByVal body As Integer) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 10/07/2008
'Checks if a given body index is a boat
'**************************************************************
'TODO : This should be checked somehow else. This is nasty....
    If body = iFragataReal Or body = iFragataCaos Or body = iBarcaPk Or _
            body = iGaleraPk Or body = iGaleonPk Or body = iBarcaCiuda Or _
            body = iGaleraCiuda Or body = iGaleonCiuda Or body = iFragataFantasmal Then
        BodyIsBoat = True
    End If
End Function

Public Sub SetInvisible(ByVal UserIndex As Integer, ByVal userCharIndex As Integer, ByVal invisible As Boolean)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim sndNick As String

With UserList(UserIndex)
    Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(userCharIndex, invisible))
    
    sndNick = .Name
    
    If invisible Then
        sndNick = sndNick & " " & TAG_USER_INVISIBLE
    Else
        If .GuildIndex > 0 Then
            sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"
        End If
    End If
    
    Call SendData(SendTarget.ToGMsAreaButRmsOrCounselors, UserIndex, PrepareMessageCharacterChangeNick(userCharIndex, sndNick))
End With
End Sub

Public Sub SetConsulatMode(ByVal UserIndex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 05/06/10
'
'***************************************************

Dim sndNick As String

With UserList(UserIndex)
    sndNick = .Name
    
    If .flags.EnConsulta Then
        sndNick = sndNick & " " & TAG_CONSULT_MODE
    Else
        If .GuildIndex > 0 Then
            sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"
        End If
    End If
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeNick(.Char.CharIndex, sndNick))
End With
End Sub

Public Function IsArena(ByVal UserIndex As Integer) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify Date: 10/11/2009
'Returns true if the user is in an Arena
'**************************************************************
    IsArena = (TriggerZonaPelea(UserIndex, UserIndex) = TRIGGER6_PERMITE)
End Function

Public Sub PerdioNpc(ByVal UserIndex As Integer, Optional ByVal CheckPets As Boolean = True)
'**************************************************************
'Author: ZaMa
'Last Modify Date: 11/07/2010 (ZaMa)
'The user loses his owned npc
'18/01/2010: ZaMa - Las mascotas dejan de atacar al npc que se perdió.
'11/07/2010: ZaMa - Coloco el indice correcto de las mascotas y ahora siguen al amo si existen.
'13/07/2010: ZaMa - Ahora solo dejan de atacar las mascotas si estan atacando al npc que pierde su amo.
'**************************************************************

    Dim PetCounter As Long
    Dim PetIndex As Integer
    Dim NpcIndex As Integer
    
    With UserList(UserIndex)
        
        NpcIndex = .flags.OwnedNpc
        If NpcIndex > 0 Then
            
            If CheckPets Then
                ' Dejan de atacar las mascotas
                If .NroMascotas > 0 Then
                    For PetCounter = 1 To MAXMASCOTAS
                    
                        PetIndex = .MascotasIndex(PetCounter)
                        
                        If PetIndex > 0 Then
                            ' Si esta atacando al npc deja de hacerlo
                            If Npclist(PetIndex).TargetNPC = NpcIndex Then
                                Call FollowAmo(PetIndex)
                            End If
                        End If
                        
                    Next PetCounter
                End If
            End If
            
            ' Reset flags
            Npclist(NpcIndex).Owner = 0
            .flags.OwnedNpc = 0

        End If
    End With
End Sub

Public Sub ApropioNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'**************************************************************
'Author: ZaMa
'Last Modify Date: 18/01/2010 (zaMa)
'The user owns a new npc
'18/01/2010: ZaMa - El sistema no aplica a zonas seguras.
'19/04/2010: ZaMa - Ahora los admins no se pueden apropiar de npcs.
'**************************************************************

    With UserList(UserIndex)
        ' Los admins no se pueden apropiar de npcs
        If EsGM(UserIndex) Then Exit Sub
        
        'No aplica a zonas seguras
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then Exit Sub
        
        ' No aplica a algunos mapas que permiten el robo de npcs
        If MapInfo(.Pos.map).RoboNpcsPermitido = 1 Then Exit Sub
        
        ' Pierde el npc anterior
        If .flags.OwnedNpc > 0 Then Npclist(.flags.OwnedNpc).Owner = 0
        
        ' Si tenia otro dueño, lo perdio aca
        Npclist(NpcIndex).Owner = UserIndex
        .flags.OwnedNpc = NpcIndex
    End With
    
    ' Inicializo o actualizo el timer de pertenencia
    Call IntervaloPerdioNpc(UserIndex, True)
End Sub

Public Function GetDireccion(ByVal UserIndex As Integer, ByVal OtherUserIndex As Integer) As String
'**************************************************************
'Author: ZaMa
'Last Modify Date: 17/11/2009
'Devuelve la direccion hacia donde esta el usuario
'**************************************************************
    Dim X As Integer
    Dim Y As Integer
    
    X = UserList(UserIndex).Pos.X - UserList(OtherUserIndex).Pos.X
    Y = UserList(UserIndex).Pos.Y - UserList(OtherUserIndex).Pos.Y
    
    If X = 0 And Y > 0 Then
        GetDireccion = "Sur"
    ElseIf X = 0 And Y < 0 Then
        GetDireccion = "Norte"
    ElseIf X > 0 And Y = 0 Then
        GetDireccion = "Este"
    ElseIf X < 0 And Y = 0 Then
        GetDireccion = "Oeste"
    ElseIf X > 0 And Y < 0 Then
        GetDireccion = "NorEste"
    ElseIf X < 0 And Y < 0 Then
        GetDireccion = "NorOeste"
    ElseIf X > 0 And Y > 0 Then
        GetDireccion = "SurEste"
    ElseIf X < 0 And Y > 0 Then
        GetDireccion = "SurOeste"
    End If

End Function

Public Function SameFaccion(ByVal UserIndex As Integer, ByVal OtherUserIndex As Integer) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify Date: 17/11/2009
'Devuelve True si son de la misma faccion
'**************************************************************
    SameFaccion = (esCaos(UserIndex) And esCaos(OtherUserIndex)) Or _
                    (esArmada(UserIndex) And esArmada(OtherUserIndex))
End Function

Public Function FarthestPet(ByVal UserIndex As Integer) As Integer
'**************************************************************
'Author: ZaMa
'Last Modify Date: 18/11/2009
'Devuelve el indice de la mascota mas lejana.
'**************************************************************
On Error GoTo Errhandler
    
    Dim PetIndex As Integer
    Dim Distancia As Integer
    Dim OtraDistancia As Integer
    
    With UserList(UserIndex)
        If .NroMascotas = 0 Then Exit Function
    
        For PetIndex = 1 To MAXMASCOTAS
            ' Solo pos invocar criaturas que exitan!
            If .MascotasIndex(PetIndex) > 0 Then
                ' Solo aplica a mascota, nada de elementales..
                If Npclist(.MascotasIndex(PetIndex)).Contadores.TiempoExistencia = 0 Then
                    If FarthestPet = 0 Then
                        ' Por si tiene 1 sola mascota
                        FarthestPet = PetIndex
                        Distancia = Abs(.Pos.X - Npclist(.MascotasIndex(PetIndex)).Pos.X) + _
                                    Abs(.Pos.Y - Npclist(.MascotasIndex(PetIndex)).Pos.Y)
                    Else
                        ' La distancia de la proxima mascota
                        OtraDistancia = Abs(.Pos.X - Npclist(.MascotasIndex(PetIndex)).Pos.X) + _
                                        Abs(.Pos.Y - Npclist(.MascotasIndex(PetIndex)).Pos.Y)
                        ' Esta mas lejos?
                        If OtraDistancia > Distancia Then
                            Distancia = OtraDistancia
                            FarthestPet = PetIndex
                        End If
                    End If
                End If
            End If
        Next PetIndex
    End With

    Exit Function
    
Errhandler:
    Call LogError("Error en FarthestPet")
End Function

''
' Set the EluSkill value at the skill.
'
' @param UserIndex  Specifies reference to user
' @param Skill      Number of the skill to check
' @param Allocation True If the motive of the modification is the allocation, False if the skill increase by training

Public Sub CheckEluSkill(ByVal UserIndex As Integer, ByVal Skill As Byte, ByVal Allocation As Boolean)
'*************************************************
'Author: Torres Patricio (Pato)
'Last modified: 11/20/2009
'
'*************************************************

With UserList(UserIndex).Stats
    If .UserSkills(Skill) < MAXSKILLPOINTS Then
        If Allocation Then
            .ExpSkills(Skill) = 0
        Else
            .ExpSkills(Skill) = .ExpSkills(Skill) - .EluSkills(Skill)
        End If
        
        .EluSkills(Skill) = ELU_SKILL_INICIAL * 1 ^ .UserSkills(Skill)
    Else
        .ExpSkills(Skill) = 0
        .EluSkills(Skill) = 0
    End If
End With

End Sub

Public Function HasEnoughItems(ByVal UserIndex As Integer, ByVal objindex As Integer, ByVal Amount As Long) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify Date: 25/11/2009
'Cheks Wether the user has the required amount of items in the inventory or not
'**************************************************************

    Dim Slot As Long
    Dim ItemInvAmount As Long
    
    For Slot = 1 To UserList(UserIndex).CurrentInventorySlots
        ' Si es el item que busco
        If UserList(UserIndex).Invent.Object(Slot).objindex = objindex Then
            ' Lo sumo a la cantidad total
            ItemInvAmount = ItemInvAmount + UserList(UserIndex).Invent.Object(Slot).Amount
        End If
    Next Slot

    HasEnoughItems = Amount <= ItemInvAmount
End Function

Public Function TotalOfferItems(ByVal objindex As Integer, ByVal UserIndex As Integer) As Long
'**************************************************************
'Author: ZaMa
'Last Modify Date: 25/11/2009
'Cheks the amount of items the user has in offerSlots.
'**************************************************************
    Dim Slot As Byte
    
    For Slot = 1 To MAX_OFFER_SLOTS
            ' Si es el item que busco
        If UserList(UserIndex).ComUsu.Objeto(Slot) = objindex Then
            ' Lo sumo a la cantidad total
            TotalOfferItems = TotalOfferItems + UserList(UserIndex).ComUsu.cant(Slot)
        End If
    Next Slot

End Function

Public Function getMaxInventorySlots(ByVal UserIndex As Integer) As Byte
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

If UserList(UserIndex).Invent.MochilaEqpObjIndex > 0 Then
    getMaxInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + ObjData(UserList(UserIndex).Invent.MochilaEqpObjIndex).MochilaType * 5 '5=slots por fila, hacer constante
Else
    getMaxInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
End If
End Function

Public Sub goHome(ByVal UserIndex As Integer)
Dim Distance As Integer
Dim tiempo As Long

With UserList(UserIndex)
    If .flags.Muerto = 1 Then
        If .flags.lastMap = 0 Then
            Distance = distanceToCities(.Pos.map).distanceToCity(.Hogar)
        Else
            Distance = distanceToCities(.flags.lastMap).distanceToCity(.Hogar) + GOHOME_PENALTY
        End If
        
        tiempo = (Distance + 1) * 30 'segundos
        
        .Counters.goHome = tiempo / 6 'Se va a chequear cada 6 segundos.
        
        .flags.Traveling = 1

        Call WriteMultiMessage(UserIndex, eMessages.Home, Distance, tiempo, , MapInfo(Ciudades(.Hogar).map).Name)
    Else
        Call WriteConsoleMsg(UserIndex, "Debes estar muerto para poder utilizar este comando.", FontTypeNames.FONTTYPE_FIGHT)
    End If
End With
End Sub

Public Function ToogleToAtackable(ByVal UserIndex As Integer, ByVal OwnerIndex As Integer, Optional ByVal StealingNpc As Boolean = True) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 15/01/2010
'Change to Atackable mode.
'***************************************************
    
    Dim AtacablePor As Integer
    
    With UserList(UserIndex)
        ' Inicializar el timer
        Call IntervaloEstadoAtacable(UserIndex, True)
        
        ToogleToAtackable = True
        
    End With
    
End Function

Public Sub setHome(ByVal UserIndex As Integer, ByVal newHome As eCiudad, ByVal NpcIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 30/04/2010
'30/04/2010: ZaMa - Ahora el npc avisa que se cambio de hogar.
'***************************************************
    If newHome < eCiudad.cUllathorpe Or newHome > cArghal Then Exit Sub
    UserList(UserIndex).Hogar = newHome
    
    Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido a nuestra humilde comunidad, este es ahora tu nuevo hogar!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
End Sub

Public Sub ToggleBoatBody(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 25/07/2010
'Gives boat body depending on user alignment.
'25/07/2010: ZaMa - Now makes difference depending on faccion and atacable status.
'***************************************************

    Dim Ropaje As Integer
    Dim EsFaccionario As Boolean
    Dim NewBody As Integer
    
    With UserList(UserIndex)
 
        .Char.Head = 0
        If .Invent.BarcoObjIndex = 0 Then Exit Sub
        
        Ropaje = ObjData(.Invent.BarcoObjIndex).Ropaje
        
        ' Criminales y caos
        If criminal(UserIndex) Then
            
            EsFaccionario = esCaos(UserIndex)
            
            Select Case Ropaje
                Case iBarca
                    If EsFaccionario Then
                        NewBody = iBarcaPk
                    Else
                        NewBody = iBarcaPk
                    End If
                
                Case iGalera
                    If EsFaccionario Then
                        NewBody = iGaleraPk
                    Else
                        NewBody = iGaleraPk
                    End If
                    
                Case iGaleon
                    If EsFaccionario Then
                        NewBody = iGaleonPk
                    Else
                        NewBody = iGaleonPk
                    End If
            End Select
        
        ' Ciudas y Armadas
        Else
            
            EsFaccionario = esArmada(UserIndex)
            
            ' Atacable
            If .flags.AtacablePor <> 0 Then
                
                Select Case Ropaje
                    Case iBarca
                        If EsFaccionario Then
                            NewBody = iBarcaCiuda
                        Else
                            NewBody = iBarcaCiuda
                        End If
                    
                    Case iGalera
                        If EsFaccionario Then
                            NewBody = iGaleraCiuda
                        Else
                            NewBody = iGaleraCiuda
                        End If
                        
                    Case iGaleon
                        If EsFaccionario Then
                            NewBody = iGaleonCiuda
                        Else
                            NewBody = iGaleonCiuda
                        End If
                End Select
            
            ' Normal
            Else
            
                Select Case Ropaje
                    Case iBarca
                        If EsFaccionario Then
                            NewBody = iBarcaCiuda
                        Else
                            NewBody = iBarcaCiuda
                        End If
                    
                    Case iGalera
                        If EsFaccionario Then
                            NewBody = iGaleraCiuda
                        Else
                            NewBody = iGaleraCiuda
                        End If
                        
                    Case iGaleon
                        If EsFaccionario Then
                            NewBody = iGaleonCiuda
                        Else
                            NewBody = iGaleonCiuda
                        End If
                End Select
            
            End If
            
        End If
        
        .Char.body = NewBody
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        .Char.CascoAnim = NingunCasco
    End With

End Sub

Sub SendUserStatsMercado(ByVal SendIndex As Integer, ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is online.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
'*************************************************
    With UserList(UserIndex)
        Call WriteConsoleMsg(SendIndex, "Nick: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Ciudadanos matados: " & .Faccion.CiudadanosMatados & " Criminales matados: " & .Faccion.CriminalesMatados & " usuarios matados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "NPCs muertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Clase: " & ListaClases(.clase), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Penas: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)
         Call WriteConsoleMsg(SendIndex, "Nivel: " & .Stats.ELV & "  EXP: " & .Stats.Exp & "/" & .Stats.ELU, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Salud: " & .Stats.MinHp & "/" & .Stats.MaxHp & "  Maná: " & .Stats.MinMAN & "/" & .Stats.MaxMAN & "  Energía: " & .Stats.MinSta & "/" & .Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
        
        If .Faccion.ArmadaReal = 1 Then
            Call WriteConsoleMsg(SendIndex, "Ejército real desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Ingresó en nivel: " & .Faccion.NivelIngreso & " con " & .Faccion.MatadosIngreso & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.FuerzasCaos = 1 Then
            Call WriteConsoleMsg(SendIndex, "Legión oscura desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Ingresó en nivel: " & .Faccion.NivelIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.RecibioExpInicialReal = 1 Then
            Call WriteConsoleMsg(SendIndex, "Fue ejército real", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.RecibioExpInicialCaos = 1 Then
            Call WriteConsoleMsg(SendIndex, "Fue legión oscura", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        End If
        
        Call WriteConsoleMsg(SendIndex, "Asesino: " & .Reputacion.AsesinoRep, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Noble: " & .Reputacion.NobleRep, FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(SendIndex, "Oro: " & .Stats.Gld & "  Posición: " & .Pos.X & "," & .Pos.Y & " en mapa " & .Pos.map, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Dados: " & .Stats.UserAtributos(eAtributos.Fuerza) & ", " & .Stats.UserAtributos(eAtributos.Agilidad) & ", " & .Stats.UserAtributos(eAtributos.Inteligencia) & ", " & .Stats.UserAtributos(eAtributos.Carisma) & ", " & .Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Retos Ganados: " & .Stats.RetosGanados & "", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Retos Perdidos: " & .Stats.RetosPerdidos & "", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Oro Ganado: " & .Stats.OroGanado & "", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Oro Perdido: " & .Stats.OroPerdido & "", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Torneos Ganados: " & .Stats.TorneosGanados & "", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Famas asignadas: " & .flags.BonosHP & "", FontTypeNames.FONTTYPE_INFO)
        
        '?
        If .GuildIndex > 0 Then
            Call WriteConsoleMsg(SendIndex, "Clan: " & GuildName(.GuildIndex), FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub
