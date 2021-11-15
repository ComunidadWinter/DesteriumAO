Attribute VB_Name = "SistemaCombate"

'Argentum Online 0.12.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'
'Dise�o y correcci�n del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

'9/01/2008 Pablo (ToxicWaste) - Ahora TODOS los modificadores de Clase se controlan desde Balance.dat


Option Explicit

Public Const MAXDISTANCIAARCO As Byte = 18
Public Const MAXDISTANCIAMAGIA As Byte = 18

Public Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        MinimoInt = b
    Else
        MinimoInt = a
    End If
End Function

Public Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        MaximoInt = a
    Else
        MaximoInt = b
    End If
End Function

Private Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(eSkill.Defensa) * ModClase(UserList(UserIndex).clase).Escudo) / 2
End Function

Private Function PoderEvasion(ByVal UserIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    Dim lTemp As Long
    With UserList(UserIndex)
        lTemp = (.Stats.UserSkills(eSkill.Tacticas) + _
          .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).Evasion
       
        PoderEvasion = (lTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        If .Stats.UserSkills(eSkill.Armas) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Armas) * ModClase(.clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Armas) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Armas) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueArmas
        Else
           PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueArmas
        End If
        
        PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        If .Stats.UserSkills(eSkill.Proyectiles) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Proyectiles) * ModClase(.clase).AtaqueProyectiles
        ElseIf .Stats.UserSkills(eSkill.Proyectiles) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueProyectiles
        ElseIf .Stats.UserSkills(eSkill.Proyectiles) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueProyectiles
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueProyectiles
        End If
        
        PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderAtaqueWrestling(ByVal UserIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        If .Stats.UserSkills(eSkill.Wrestling) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Wrestling) * ModClase(.clase).AtaqueWrestling
        ElseIf .Stats.UserSkills(eSkill.Wrestling) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueWrestling
        ElseIf .Stats.UserSkills(eSkill.Wrestling) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueWrestling
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueWrestling
        End If
        
        PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Public Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim PoderAtaque As Long
    Dim Arma As Integer
    Dim Skill As eSkill
    Dim ProbExito As Long
    
    Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
    
    If Arma > 0 Then 'Usando un arma
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(UserIndex)
            Skill = eSkill.Proyectiles
        Else
            PoderAtaque = PoderAtaqueArma(UserIndex)
            Skill = eSkill.Armas
        End If
    Else 'Peleando con pu�os
        PoderAtaque = PoderAtaqueWrestling(UserIndex)
        Skill = eSkill.Wrestling
    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))
    
    UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
    
    If UserImpactoNpc Then
        Call SubirSkill(UserIndex, Skill, True)
    Else
        Call SubirSkill(UserIndex, Skill, False)
    End If
End Function
Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
    On Error GoTo NpcImpacto_Error
      '*************************************************
      'Author: Unknown
      'Last modified: 03/15/2006
      'Revisa si un NPC logra impactar a un user o no
      '03/15/2006 Maraxus - Evit� una divisi�n por cero que eliminaba NPCs
      '*************************************************
          Dim Rechazo As Boolean
          Dim ProbRechazo As Long
          Dim ProbExito As Long
          Dim UserEvasion As Long
          Dim NpcPoderAtaque As Long
          Dim PoderEvasioEscudo As Long
          Dim SkillTacticas As Long
          Dim SkillDefensa As Long
          
'Agregu� un Manejador de error que me dice _
en qu� linea va a tirar el Error.

10        UserEvasion = PoderEvasion(UserIndex)
20        NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
30        PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)
          
40        SkillTacticas = UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas)
50        SkillDefensa = UserList(UserIndex).Stats.UserSkills(eSkill.Defensa)
          
          'Esta usando un escudo ???
60        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo
          
          ' Chances are rounded
70        ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))
          
80        NpcImpacto = (RandomNumber(1, 100) <= ProbExito)
          
          ' el usuario esta usando un escudo ???
90        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
100           If Not NpcImpacto Then
110               If SkillDefensa + SkillTacticas > 0 Then  'Evitamos divisi�n por cero
                      ' Chances are rounded
120                   ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
130                   Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                      
140                   If Rechazo Then
                          'Se rechazo el ataque con el escudo
150                       Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
160                       Call WriteMultiMessage(UserIndex, eMessages.BlockedWithShieldUser) 'Call WriteBlockedWithShieldUser(UserIndex)
                        Call WriteChatOverHead(UserIndex, "�Bloqueo!", UserList(UserIndex).Char.CharIndex, vbYellow)
170                       SendData SendTarget.ToPCArea, UserIndex, PrepareMessageMovimientSW(UserList(UserIndex).Char.CharIndex, 2)
180                       Call SubirSkill(UserIndex, eSkill.Defensa, True)
190                   Else
200                       Call SubirSkill(UserIndex, eSkill.Defensa, False)
210                   End If
220               End If
230           End If
240       End If
    
    On Error GoTo 0
    Exit Function

NpcImpacto_Error:

     LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure NpcImpacto, line " & Erl & "."
    
End Function




Public Function CalcularDa�o(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
'***************************************************
'Author: Unknown
'Last Modification: 01/04/2010 (ZaMa)
'01/04/2010: ZaMa - Modifico el da�o de wrestling.
'01/04/2010: ZaMa - Agrego bonificadores de wrestling para los guantes.
'***************************************************
    Dim Da�oArma As Long
    Dim Da�oUsuario As Long
    Dim Arma As ObjData
    Dim ModifClase As Single
    Dim proyectil As ObjData
    Dim Da�oMaxArma As Long
    Dim Da�oMinArma As Long
    Dim objindex As Integer
    
    ''sacar esto si no queremos q la matadracos mate el Dragon si o si
    Dim matoDragon As Boolean
    matoDragon = False
    
    With UserList(UserIndex)
        If .Invent.WeaponEqpObjIndex > 0 Then
            Arma = ObjData(.Invent.WeaponEqpObjIndex)
            
            ' Ataca a un npc?
            If NpcIndex > 0 Then
                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.clase).Da�oProyectiles
                    Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    Da�oMaxArma = Arma.MaxHIT
                    
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        Da�oArma = Da�oArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                        ' For some reason this isn't done...
                        'Da�oMaxArma = Da�oMaxArma + proyectil.MaxHIT
                    End If
                Else
                    ModifClase = ModClase(.clase).Da�oArmas
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then ' Usa la mata Dragones?
                        If Npclist(NpcIndex).NPCtype = Dragon Then 'Ataca Dragon?
                            Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                            Da�oMaxArma = Arma.MaxHIT
                             matoDragon = True
                        Else ' Sino es Dragon da�o es 1
                            Da�oArma = 1
                            Da�oMaxArma = 1
                        End If
                    Else
                        Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                        Da�oMaxArma = Arma.MaxHIT
                    End If
                End If
            Else ' Ataca usuario
                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.clase).Da�oProyectiles
                    Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    Da�oMaxArma = Arma.MaxHIT
                     
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        Da�oArma = Da�oArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                        ' For some reason this isn't done...
                        'Da�oMaxArma = Da�oMaxArma + proyectil.MaxHIT
                    End If
                Else
                    ModifClase = ModClase(.clase).Da�oArmas
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                        ModifClase = ModClase(.clase).Da�oArmas
                        Da�oArma = 1 ' Si usa la espada mataDragones da�o es 1
                        Da�oMaxArma = 1
                    Else
                        Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                        Da�oMaxArma = Arma.MaxHIT
                    End If
                End If
            End If
        Else
            ModifClase = ModClase(.clase).Da�oWrestling
            
            ' Da�o sin guantes
            Da�oMinArma = 4
            Da�oMaxArma = 9
            
            ' Plus de guantes (en slot de anillo)
            objindex = .Invent.AnilloEqpObjIndex
            If objindex > 0 Then
                If ObjData(objindex).Guante = 1 Then
                    Da�oMinArma = Da�oMinArma + ObjData(objindex).MinHIT
                    Da�oMaxArma = Da�oMaxArma + ObjData(objindex).MaxHIT
                End If
            End If
            
            Da�oArma = RandomNumber(Da�oMinArma, Da�oMaxArma)
            
        End If
        
        Da�oUsuario = RandomNumber(.Stats.MinHIT, .Stats.MaxHIT)
        
        ''sacar esto si no queremos q la matadracos mate el Dragon si o si
        If matoDragon Then
            CalcularDa�o = Npclist(NpcIndex).Stats.MinHp + Npclist(NpcIndex).Stats.def
        Else
            CalcularDa�o = (3 * Da�oArma + ((Da�oMaxArma / 5) * MaximoInt(0, .Stats.UserAtributos(eAtributos.Fuerza) - 15)) + Da�oUsuario) * ModifClase
        End If
    End With
End Function

Public Sub UserDa�oNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    On Error GoTo UserDa�oNpc_Error
      '***************************************************
      'Author: Unknown
      'Last Modification: 07/04/2010 (ZaMa)
      '25/01/2010: ZaMa - Agrego poder acuchillar npcs.
      '07/04/2010: ZaMa - Los asesinos apu�alan acorde al da�o base sin descontar la defensa del npc.
      '***************************************************

          Dim da�o As Long
          Dim Da�oBase As Long
          
10        Da�oBase = CalcularDa�o(UserIndex, NpcIndex)
          
          'esta navegando? si es asi le sumamos el da�o del barco
20        If UserList(UserIndex).flags.Navegando = 1 Then
30            If UserList(UserIndex).Invent.BarcoObjIndex > 0 Then
40                Da�oBase = Da�oBase + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHIT, _
                                              ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MaxHIT)
50            End If
60        End If
          
          
70        With Npclist(NpcIndex)
80            da�o = Da�oBase - .Stats.def
              
            ' EL QUE TE PEGA MAMU al NPC
            If StrComp(GreatPower.CurrentUser, UCase$(UserList(UserIndex).Name)) = 0 Then
                da�o = da�o * 1.3
            End If
        
90            If da�o < 0 Then da�o = 0
              
              'Call WriteUserHitNPC(UserIndex, da�o)
100           Call WriteMultiMessage(UserIndex, eMessages.UserHitNPC, da�o)
110           Call CalcularDarExp(UserIndex, NpcIndex, da�o)
120           .Stats.MinHp = .Stats.MinHp - da�o
              
130           SendData SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, da�o, DAMAGE_NORMAL)
                     
140           If .Stats.MinHp > 0 Then
                  'Trata de apu�alar por la espalda al enemigo
150               If UserList(UserIndex).clase = eClass.Hunter Then
160               DoGolpeArco UserIndex, NpcIndex, 0, da�o
170               End If
                  
180               If PuedeApu�alar(UserIndex) Then
190               UserList(UserIndex).Da�oapu = da�o
200                  Call DoApu�alar(UserIndex, NpcIndex, 0, Da�oBase)
210               End If
                  
                  'trata de dar golpe cr�tico
220               Call DoGolpeCritico(UserIndex, NpcIndex, 0, da�o)
                  
230               If PuedeAcuchillar(UserIndex) Then
240                   Call DoAcuchillar(UserIndex, NpcIndex, 0, da�o)
250               End If
260           End If
              
              
270           If .Stats.MinHp <= 0 Then
                  ' Si era un Dragon perdemos la espada mataDragones
280               If .NPCtype = Dragon Then
                      'Si tiene equipada la matadracos se la sacamos
290                   If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
300                       Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)
310                   End If
320                   If .Stats.MaxHp > 100000 Then Call LogDesarrollo(UserList(UserIndex).Name & " mat� un drag�n")
330               End If
                  
                  ' Para que las mascotas no sigan intentando luchar y
                  ' comiencen a seguir al amo
                  Dim j As Integer
340               For j = 1 To MAXMASCOTAS
350                   If UserList(UserIndex).MascotasIndex(j) > 0 Then
360                       If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex Then
370                           Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0
380                           Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.SigueAmo
390                       End If
400                   End If
410               Next j
                  
420               Call MuereNpc(NpcIndex, UserIndex)
430           End If
440       End With
    
    On Error GoTo 0
    Exit Sub

UserDa�oNpc_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure UserDa�oNpc, line " & Erl & "."
    
End Sub



Public Sub NpcDa�o(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
 
    Dim da�o As Integer
    Dim Lugar As Integer
    Dim absorbido As Integer
    Dim defbarco As Integer
    Dim defmontura As Integer
    Dim Obj As ObjData
   
    da�o = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHIT)
    
    With UserList(UserIndex)
        If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
            Obj = ObjData(.Invent.BarcoObjIndex)
            defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
       
        If .flags.Montando = 1 Then
            Obj = ObjData(.Invent.MonturaObjIndex)
            defmontura = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
       
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
       
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                'Si tiene casco absorbe el golpe
                If .Invent.CascoEqpObjIndex > 0 Then
                   Obj = ObjData(.Invent.CascoEqpObjIndex)
                   absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                End If
          Case Else
                'Si tiene armadura absorbe el golpe
                If .Invent.ArmourEqpObjIndex > 0 Then
                    Dim Obj2 As ObjData
                    Obj = ObjData(.Invent.ArmourEqpObjIndex)
                    If .Invent.EscudoEqpObjIndex Then
                        Obj2 = ObjData(.Invent.EscudoEqpObjIndex)
                        absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
                    Else
                        absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                   End If
                End If
        End Select
       
        absorbido = absorbido + defbarco + defmontura
        da�o = da�o - absorbido
        If da�o < 1 Then da�o = 1
       
        Call WriteMultiMessage(UserIndex, eMessages.NPCHitUser, Lugar, da�o)
        'Call WriteNPCHitUser(UserIndex, Lugar, da�o)
       
        If .flags.Privilegios And PlayerType.User Then .Stats.MinHp = .Stats.MinHp - da�o
       
        If .flags.Meditando Then
            If da�o > Fix(.Stats.MinHp / 100 * .Stats.UserAtributos(eAtributos.Inteligencia) * .Stats.UserSkills(eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
                .flags.Meditando = False
                Call WriteMeditateToggle(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                .Char.FX = 0
                .Char.loops = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            End If
        End If
       
        'Muere el usuario
        If .Stats.MinHp <= 0 Then
            Call WriteMultiMessage(UserIndex, eMessages.NPCKillUser) 'Call WriteNPCKillUser(UserIndex) ' Le informamos que ha muerto
           
            'Si lo mato un guardia
            If criminal(UserIndex) And Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                Call RestarCriminalidad(UserIndex)
                If Not criminal(UserIndex) And .Faccion.FuerzasCaos = 1 Then Call ExpulsarFaccionCaos(UserIndex)
            End If
           
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
            Else
                'Al matarlo no lo sigue mas
                If Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
                    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
                    Npclist(NpcIndex).flags.AttackedBy = vbNullString
                End If
            End If
           
            Call UserDie(UserIndex)
        End If
    End With
End Sub


Public Sub RestarCriminalidad(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim EraCriminal As Boolean
    EraCriminal = criminal(UserIndex)
    
    With UserList(UserIndex).Reputacion
        If .BandidoRep > 0 Then
             .BandidoRep = .BandidoRep - vlASALTO
             If .BandidoRep < 0 Then .BandidoRep = 0
        ElseIf .LadronesRep > 0 Then
             .LadronesRep = .LadronesRep - (vlCAZADOR * 10)
             If .LadronesRep < 0 Then .LadronesRep = 0
        End If
    End With
    
    If EraCriminal And Not criminal(UserIndex) Then
        Call RefreshCharStatus(UserIndex)
    End If
End Sub

Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, Optional ByVal CheckElementales As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: 15/04/2010
'15/04/2010: ZaMa - Las mascotas no se apropian de npcs.
'***************************************************

    Dim j As Integer
    
    ' Si no tengo mascotas, para que cheaquear lo demas?
    If UserList(UserIndex).NroMascotas = 0 Then Exit Sub
    
    If Not PuedeAtacarNPC(UserIndex, NpcIndex, , True) Then Exit Sub
    
    With UserList(UserIndex)
        For j = 1 To MAXMASCOTAS
            If .MascotasIndex(j) > 0 Then
               If .MascotasIndex(j) <> NpcIndex Then
                If CheckElementales Or (Npclist(.MascotasIndex(j)).Numero <> ELEMENTALFUEGO And Npclist(.MascotasIndex(j)).Numero <> ELEMENTALTIERRA) Then
                    
                    If Npclist(.MascotasIndex(j)).TargetNPC = 0 Then Npclist(.MascotasIndex(j)).TargetNPC = NpcIndex
                    Npclist(.MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc
                End If
               End If
            End If
        Next j
    End With
End Sub

Public Sub AllFollowAmo(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim j As Integer
    
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(j) > 0 Then
            Call FollowAmo(UserList(UserIndex).MascotasIndex(j))
        End If
    Next j
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
'*************************************************
'Author: Unknown
'Last modified: -
'
'*************************************************

    With UserList(UserIndex)
        If .flags.AdminInvisible = 1 Then Exit Function
        If (Not .flags.Privilegios And PlayerType.User) <> 0 And Not .flags.AdminPerseguible Then Exit Function
        If Not CanAttackReyCastle(UserIndex, NpcIndex) Then Exit Function
    End With
    
    With Npclist(NpcIndex)
        ' El npc puede atacar ???
        If .CanAttack = 1 Then
            NpcAtacaUser = True
            Call CheckPets(NpcIndex, UserIndex, False)
            
            If .Target = 0 Then .Target = UserIndex
            
            If UserList(UserIndex).flags.AtacadoPorNpc = 0 And UserList(UserIndex).flags.AtacadoPorUser = 0 Then
                UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex
            End If
        Else
            NpcAtacaUser = False
            Exit Function
        End If
        
        .CanAttack = 0
        
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))
        End If
    End With
    
    If NpcImpacto(NpcIndex, UserIndex) Then
        With UserList(UserIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            
            If .flags.Meditando = False Then
                If .flags.Navegando = 0 And .flags.Montando = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0))
                End If
            End If
            
            Call NpcDa�o(NpcIndex, UserIndex)
            Call WriteUpdateHP(UserIndex)
            Call WriteUpdateFollow(UserIndex)
            
            '�Puede envenenar?
            If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(UserIndex)
        End With
        
        Call SubirSkill(UserIndex, eSkill.Tacticas, False)
    Else
        Call WriteMultiMessage(UserIndex, eMessages.NPCSwing)
        Call SubirSkill(UserIndex, eSkill.Tacticas, True)
    End If
    
    'Controla el nivel del usuario
    Call CheckUserLevel(UserIndex)
End Function

Private Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim PoderAtt As Long
    Dim PoderEva As Long
    Dim ProbExito As Long
    
    PoderAtt = Npclist(Atacante).PoderAtaque
    PoderEva = Npclist(Victima).PoderEvasion
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtt - PoderEva) * 0.4))
    NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
End Function

Public Sub NpcDa�oNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim da�o As Integer
    
    With Npclist(Atacante)
        da�o = RandomNumber(.Stats.MinHIT, .Stats.MaxHIT)
        Npclist(Victima).Stats.MinHp = Npclist(Victima).Stats.MinHp - da�o
        CalcularDarExp .MaestroUser, Victima, da�o
        If Npclist(Victima).Stats.MinHp < 1 Then
            .Movement = .flags.OldMovement
            
            If LenB(.flags.AttackedBy) <> 0 Then
                .Hostile = .flags.OldHostil
            End If
            
            If .MaestroUser > 0 Then
                Call FollowAmo(Atacante)
            End If
            
            Call MuereNpc(Victima, .MaestroUser)
        End If
    End With
End Sub
        

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)
'*************************************************
'Author: Unknown
'Last modified: 01/03/2009
'01/03/2009: ZaMa - Las mascotas no pueden atacar al rey si quedan pretorianos vivos.
'23/05/2010: ZaMa - Ahora los elementales renuevan el tiempo de pertencia del npc que atacan si pertenece a su amo.
'*************************************************
   
    Dim MasterIndex As Integer
   
    With Npclist(Atacante)
       
         'Es el Rey Preatoriano?
        If Npclist(Victima).Numero = PRKING_NPC Then
            If pretorianosVivos > 0 Then
                Call WriteConsoleMsg(.MaestroUser, "Debes matar al resto del ej�rcito antes de atacar al rey!", FontTypeNames.FONTTYPE_FIGHT)
                .TargetNPC = 0
                Exit Sub
            End If
        End If
       
        ' El npc puede atacar ???
        If .CanAttack = 1 Then
            .CanAttack = 0
            If cambiarMOvimiento Then
                Npclist(Victima).TargetNPC = Atacante
                Npclist(Victima).Movement = TipoAI.NpcAtacaNpc
            End If
        Else
            Exit Sub
        End If
       
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))
        End If
       
        MasterIndex = .MaestroUser
       
        ' Tiene maestro?
        If MasterIndex > 0 Then
            ' Su maestro es due�o del npc al que ataca?
            If Npclist(Victima).Owner = MasterIndex Then
                ' Renuevo el timer de pertenencia
                Call IntervaloPerdioNpc(MasterIndex, True)
            End If
        End If
       
       
        If NpcImpactoNpc(Atacante, Victima) Then
            If Npclist(Victima).flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            End If
       
            If MasterIndex > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            End If
           
            Call NpcDa�oNpc(Atacante, Victima)
        Else
            If MasterIndex > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_SWING, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            End If
        End If
    End With
End Sub
Public Function UsuarioAtacaNpc(ByVal UserIndex As Integer, _
                                ByVal NpcIndex As Integer) As Boolean
 
        '***************************************************
        'Author: Unknown
        'Last Modification: 14/01/2010 (ZaMa)
        '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados por npcs cuando los atacan.
        '14/01/2010: ZaMa - Lo transformo en funci�n, para que no se pierdan municiones al atacar targets inv�lidos.
        '***************************************************
        '<EhHeader>
        On Error GoTo UsuarioAtacaNpc_Err
 
        '</EhHeader>
 
100     If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then Exit Function
 
102     Call NPCAtacado(NpcIndex, UserIndex)
 
104     If UserImpactoNpc(UserIndex, NpcIndex) Then
106         If Npclist(NpcIndex).flags.Snd2 > 0 Then
108             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, _
                        Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
            Else
110             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2, Npclist( _
                        NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
 
            End If
     
112         Call UserDa�oNpc(UserIndex, NpcIndex)
        Else
114         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, UserList( _
                    UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
116         Call WriteMultiMessage(UserIndex, eMessages.UserSwing)
            Call WriteChatOverHead(UserIndex, "�Fall�!", UserList(UserIndex).Char.CharIndex, vbYellow)
 
        End If
 
        ' Revel� su condici�n de usuario al atacar, los npcs lo van a atacar
118     UserList(UserIndex).flags.Ignorado = False
 
120     UsuarioAtacaNpc = True
 
        '<EhFooter>
        Exit Function
 
UsuarioAtacaNpc_Err:
        LogError Err.Description & vbCrLf & "UsuarioAtacaNpc " & NpcIndex & " " & "at line " & Erl
             
        '</EhFooter>
End Function
Public Sub UsuarioAtaca(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim index As Integer
    Dim AttackPos As WorldPos
    
    'Check bow's interval
    If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
    
    'Check Spell-Magic interval
    If Not IntervaloPermiteMagiaGolpe(UserIndex) Then
        'Check Attack interval
        If Not IntervaloPermiteAtacar(UserIndex) Then
            Exit Sub
        End If
    End If
    
        Dim loquebaja As Byte
    
    With UserList(UserIndex)
   If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).QuitaEnergia = 0 Then
            loquebaja = RandomNumber(1, 10)
                If .Stats.MinSta - loquebaja <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Est�s muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                Else
                Call QuitarSta(UserIndex, loquebaja)
                End If
            Else
                If UserList(UserIndex).Stats.MinSta >= ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).QuitaEnergia Then
                Call QuitarSta(UserIndex, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).QuitaEnergia)
                Else
                Call WriteConsoleMsg(UserIndex, "Est�s muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                End If
            End If
        Else
        Call QuitarSta(UserIndex, RandomNumber(1, 10))
        End If
        SendData SendTarget.ToPCArea, UserIndex, PrepareMessageMovimientSW(.Char.CharIndex, 1)
        
        AttackPos = .Pos
        Call HeadtoPos(.Char.Heading, AttackPos)

        'Exit if not legal
        If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Exit Sub
        End If
        
        index = MapData(AttackPos.map, AttackPos.X, AttackPos.Y).UserIndex
        
        'Look for user
        If index > 0 Then
            Call UsuarioAtacaUsuario(UserIndex, index)
            Call WriteUpdateUserStats(UserIndex)
            Call WriteUpdateUserStats(index)
            Exit Sub
        End If
        
        index = MapData(AttackPos.map, AttackPos.X, AttackPos.Y).NpcIndex
        
       'Look for NPC
        If index > 0 Then
            If Npclist(index).Attackable Then
                If Npclist(index).MaestroUser > 0 And MapInfo(Npclist(index).Pos.map).Pk = False Then
                    Call WriteConsoleMsg(UserIndex, "No puedes atacar mascotas en zona segura.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
               
                Call UsuarioAtacaNpc(UserIndex, index)
            Else
                Call WriteConsoleMsg(UserIndex, "No puedes atacar a este NPC.", FontTypeNames.FONTTYPE_FIGHT)
            End If
           
            Call WriteUpdateUserStats(UserIndex)
           
            Exit Sub
        End If
       
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
        Call WriteUpdateUserStats(UserIndex)
       
        If .Counters.Trabajando Then .Counters.Trabajando = .Counters.Trabajando - 1
           
        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
    End With
End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
    On Error GoTo UsuarioImpacto_Error
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim ProbRechazo As Long
          Dim Rechazo As Boolean
          Dim ProbExito As Long
          Dim PoderAtaque As Long
          Dim UserPoderEvasion As Long
          Dim UserPoderEvasionEscudo As Long
          Dim Arma As Integer
          Dim SkillTacticas As Long
          Dim SkillDefensa As Long
          Dim ProbEvadir As Long
          Dim Skill As eSkill
          
10        SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(eSkill.Tacticas)
20        SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(eSkill.Defensa)
          
30        Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
          
          'Calculamos el poder de evasion...
40        UserPoderEvasion = PoderEvasion(VictimaIndex)
          
50        If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
60           UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
70           UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
80        Else
90            UserPoderEvasionEscudo = 0
100       End If
          
          'Esta usando un arma ???
110       If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
120           If ObjData(Arma).proyectil = 1 Then
130               PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
140               Skill = eSkill.Proyectiles
150           Else
160               PoderAtaque = PoderAtaqueArma(AtacanteIndex)
170               Skill = eSkill.Armas
180           End If
190       Else
200           PoderAtaque = PoderAtaqueWrestling(AtacanteIndex)
210           Skill = eSkill.Wrestling
220       End If
          
          ' Chances are rounded
230       ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtaque - UserPoderEvasion) * 0.4))
          
          ' Se reduce la evasion un 25%
240       If UserList(VictimaIndex).flags.Meditando = True Then
250           ProbEvadir = (100 - ProbExito) * 0.75
260           ProbExito = MinimoInt(90, 100 - ProbEvadir)
270       End If
          
280       UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
          
          ' el usuario esta usando un escudo ???
290       If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
              'Fallo ???
300           If Not UsuarioImpacto Then
                  ' Chances are rounded
310               ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
320               Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
330               If Rechazo Then
                      'Se rechazo el ataque con el escudo
340                   Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y))
350                                    SendData SendTarget.ToPCArea, VictimaIndex, PrepareMessageMovimientSW(UserList(VictimaIndex).Char.CharIndex, 2)
360                   Call WriteMultiMessage(AtacanteIndex, eMessages.BlockedWithShieldother)
                        Call WriteChatOverHead(VictimaIndex, "�Bloqueo!", UserList(VictimaIndex).Char.CharIndex, vbYellow)
370                   Call WriteMultiMessage(VictimaIndex, eMessages.BlockedWithShieldUser)
                      
380                   Call SubirSkill(VictimaIndex, eSkill.Defensa, True)
390               Else
400                   Call SubirSkill(VictimaIndex, eSkill.Defensa, False)
410               End If
420           End If
430       End If
          
440       If Not UsuarioImpacto Then
450           Call SubirSkill(AtacanteIndex, Skill, False)
460       End If
          
470       Call FlushBuffer(VictimaIndex)
          
480       Exit Function
          
    
    On Error GoTo 0
    Exit Function

UsuarioImpacto_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure UsuarioImpacto, line " & Erl & "."

End Function

Public Function UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Lo transformo en funci�n, para que no se pierdan municiones al atacar targets
'                    inv�lidos, y evitar un doble chequeo innecesario
'***************************************************

On Error GoTo Errhandler

    If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Function
    With UserList(AtacanteIndex)
    
              
                      '  If .death = True And DeathMatch.Cuenta > 0 Then
            'WriteConsoleMsg AtacanteIndex, "�No puedes atacar antes de la cuenta regresiva!", FontTypeNames.FONTTYPE_INFO
           ' Exit Function
       ' End If
        
    'If (DeathMatch.Ingresaron < DeathMatch.Cupos And .death = True) Then
   ' WriteConsoleMsg AtacanteIndex, "No puedes atacar si no se llenaron los cupos", FontTypeNames.FONTTYPE_WARNING
    'Exit Function
    'End If
    
                    If .hungry = True And JDH.Cuenta > 0 Then
            WriteConsoleMsg AtacanteIndex, "�No puedes atacar antes de la cuenta regresiva!", FontTypeNames.FONTTYPE_INFO
            Exit Function
        End If
        
    If (JDH.Ingresaron < JDH.Cupos And .hungry = True) Then
    WriteConsoleMsg AtacanteIndex, "No puedes atacar si no se llenaron los cupos", FontTypeNames.FONTTYPE_WARNING
    Exit Function
    End If
    
        If Distancia(.Pos, UserList(VictimaIndex).Pos) > MAXDISTANCIAARCO Then
           Call WriteConsoleMsg(AtacanteIndex, "Est�s muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
           Exit Function
        End If
        
        Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)
        
        If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            
            If UserList(VictimaIndex).flags.Navegando = 0 Then
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, FXSANGRE, 0))
            End If
            
            'Pablo (ToxicWaste): Guantes de Hurto del Bandido en acci�n
                
            'y ahora, el ladr�n puede llegar a paralizar con el golpe.
            If .clase = eClass.Thief Then
                Call DoHandInmo(AtacanteIndex, VictimaIndex)
            End If
            
            Call SubirSkill(VictimaIndex, eSkill.Tacticas, False)
            Call UserDa�oUser(AtacanteIndex, VictimaIndex)
        Else
            ' Invisible admins doesn't make sound to other clients except itself
            If .flags.AdminInvisible = 1 Then
                Call EnviarDatosASlot(AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            End If
            
            Call WriteMultiMessage(AtacanteIndex, eMessages.UserSwing)
            Call WriteChatOverHead(AtacanteIndex, "�Fall�!", UserList(AtacanteIndex).Char.CharIndex, vbYellow)
            Call WriteMultiMessage(VictimaIndex, eMessages.UserAttackedSwing, AtacanteIndex)
            Call SubirSkill(VictimaIndex, eSkill.Tacticas, True)
        End If
        
        If .clase = eClass.Thief Then Call Desarmar(AtacanteIndex, VictimaIndex)
        
    End With
    
    UsuarioAtacaUsuario = True
    
    Exit Function
    
Errhandler:
    Call LogError("Error en UsuarioAtacaUsuario. Error " & Err.Number & " : " & Err.Description)
End Function

Public Sub UserDa�oUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2010 (ZaMa)
'12/01/2010: ZaMa - Implemento armas arrojadizas y probabilidad de acuchillar
'11/03/2010: ZaMa - Ahora no cuenta la muerte si estaba en estado atacable, y no se vuelve criminal
'***************************************************
   
On Error GoTo Errhandler
 
    Dim da�o As Long
    Dim Lugar As Byte
    Dim absorbido As Long
    Dim defbarco As Integer
    Dim defmontura As Integer
    Dim Obj As ObjData
    Dim Resist As Byte
   
    da�o = CalcularDa�o(AtacanteIndex)
   
    Call UserEnvenena(AtacanteIndex, VictimaIndex)
   
    With UserList(AtacanteIndex)
        If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
             Obj = ObjData(.Invent.BarcoObjIndex)
             da�o = da�o + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
        End If
        
        If .flags.Montando = 1 Then
             Obj = ObjData(.Invent.MonturaObjIndex)
             da�o = da�o + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
        End If
        
        ' PRIMER CASO: EL QUE TE PEGA MAMU TIENE EL PODER
        If StrComp(GreatPower.CurrentUser, UCase$(.Name)) = 0 Then
            da�o = da�o * 1.1
        End If
        
        ' SEGUNDO CASO: VICTIMA TIENE EL PODER
        If StrComp(GreatPower.CurrentUser, UCase$(UserList(VictimaIndex).Name)) = 0 Then
            da�o = da�o * 0.9
        End If
       
        If UserList(VictimaIndex).flags.Navegando = 1 And UserList(VictimaIndex).Invent.BarcoObjIndex > 0 Then
             Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
             defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        If UserList(VictimaIndex).flags.Montando = 1 And UserList(VictimaIndex).Invent.MonturaObjIndex > 0 Then
             Obj = ObjData(UserList(VictimaIndex).Invent.MonturaObjIndex)
             defmontura = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
       
        If .Invent.WeaponEqpObjIndex > 0 Then
            Resist = ObjData(.Invent.WeaponEqpObjIndex).Refuerzo
        End If
       
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
       
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                'Si tiene casco absorbe el golpe
                If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
                    absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    absorbido = absorbido + defbarco + defmontura - Resist
                    da�o = da�o - absorbido
                    If da�o < 0 Then da�o = 1
                End If
           
            Case Else
                'Si tiene armadura absorbe el golpe
                If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
                    Dim Obj2 As ObjData
                    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then
                        Obj2 = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
                        absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
                    Else
                        absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    End If
                    absorbido = absorbido + defbarco + defmontura - Resist
                    da�o = da�o - absorbido
                    If da�o < 0 Then da�o = 1
                End If
        End Select
        Call WriteMultiMessage(AtacanteIndex, eMessages.UserHittedUser, UserList(VictimaIndex).Char.CharIndex, Lugar, da�o)
        Call WriteMultiMessage(VictimaIndex, eMessages.UserHittedByUser, .Char.CharIndex, Lugar, da�o)
        UserList(VictimaIndex).Stats.MinHp = UserList(VictimaIndex).Stats.MinHp - da�o
        
        SendData SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateDamage(UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y, da�o, DAMAGE_NORMAL)
        
        If .flags.Hambre = 0 And .flags.Sed = 0 Then
            'Si usa un arma quizas suba "Combate con armas"
            If .Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(.Invent.WeaponEqpObjIndex).proyectil Then
                    'es un Arco. Sube Armas a Distancia
                    Call SubirSkill(AtacanteIndex, eSkill.Proyectiles, True)
                   
                    ' Si es arma arrojadiza..
                    If ObjData(.Invent.WeaponEqpObjIndex).Municion = 0 Then
                        ' Si acuchilla
                        If ObjData(.Invent.WeaponEqpObjIndex).Acuchilla = 1 Then
                            Call DoAcuchillar(AtacanteIndex, 0, VictimaIndex, da�o)
                        End If
                    End If
                Else
                    'Sube combate con armas.
                    Call SubirSkill(AtacanteIndex, eSkill.Armas, True)
                End If
            Else
                'sino tal vez lucha libre
                Call SubirSkill(AtacanteIndex, eSkill.Wrestling, True)
            End If
                   
            'Trata de apu�alar por la espalda al enemigo
             If PuedeApu�alar(AtacanteIndex) Then
            UserList(AtacanteIndex).Da�oapu = da�o
                Call DoApu�alar(AtacanteIndex, 0, VictimaIndex, da�o)
            End If
            'e intenta dar un golpe cr�tico [Pablo (ToxicWaste)]
                        Call DoGolpeCritico(AtacanteIndex, 0, VictimaIndex, da�o)
        End If
            
                   'Doble chekeo innecesario, pero bueno..
        'Hecho para que no envie apu + golpe normal.
        If Not PuedeApu�alar(AtacanteIndex) Then
               SendData SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateDamage(UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y, da�o, DAMAGE_NORMAL)
        End If
        
        If UserList(VictimaIndex).Stats.MinHp <= 0 Then
           
            ' No cuenta la muerte si estaba en estado atacable
            If UserList(VictimaIndex).flags.AtacablePor <> AtacanteIndex Then
                'Store it!
                Call Statistics.StoreFrag(AtacanteIndex, VictimaIndex)
                
                Call ContarMuerte(VictimaIndex, AtacanteIndex)
            End If
            
            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            Dim j As Integer
            For j = 1 To MAXMASCOTAS
                If .MascotasIndex(j) > 0 Then
                    If Npclist(.MascotasIndex(j)).Target = VictimaIndex Then
                        Npclist(.MascotasIndex(j)).Target = 0
                        Call FollowAmo(.MascotasIndex(j))
                    End If
                End If
            Next j
           
            Call ActStats(VictimaIndex, AtacanteIndex)
            Call UserDie(VictimaIndex, AtacanteIndex)
        Else
            'Est� vivo - Actualizamos el HP
            Call WriteUpdateHP(VictimaIndex)
            Call WriteUpdateFollow(VictimaIndex)
        End If
    End With
   
    'Controla el nivel del usuario
    Call CheckUserLevel(AtacanteIndex)
   
    Call FlushBuffer(VictimaIndex)
   
    Exit Sub
   
Errhandler:
    Dim AtacanteNick As String
    Dim VictimaNick As String
   
    If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
    If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
   
    Call LogError("Error en UserDa�oUser. Error " & Err.Number & " : " & Err.Description & " AtacanteIndex: " & _
             AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick)
End Sub

Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal victimIndex As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 05/05/2010
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
'05/05/2010: ZaMa - Ahora no suma puntos de bandido al atacar a alguien en estado atacable.
'***************************************************

    If TriggerZonaPelea(AttackerIndex, victimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
    Dim EraCriminal As Boolean
    Dim VictimaEsAtacable As Boolean
    
    If Not criminal(AttackerIndex) Then
        If Not criminal(victimIndex) Then
            ' Si la victima no es atacable por el agresor, entonces se hace pk
            VictimaEsAtacable = UserList(victimIndex).flags.AtacablePor = AttackerIndex
            If Not VictimaEsAtacable Then Call VolverCriminal(AttackerIndex)
        End If
    End If
    
    With UserList(victimIndex)
        If .flags.Meditando Then
            .flags.Meditando = False
            Call WriteMeditateToggle(victimIndex)
            Call WriteConsoleMsg(victimIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, victimIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
        End If
    End With
    
    EraCriminal = criminal(AttackerIndex)
    
    ' Si ataco a un atacable, no suma puntos de bandido
    If Not VictimaEsAtacable Then
        With UserList(AttackerIndex).Reputacion
            If Not criminal(victimIndex) Then
                .BandidoRep = .BandidoRep + vlASALTO
                If .BandidoRep > MAXREP Then .BandidoRep = MAXREP
                
                .NobleRep = .NobleRep * 0.5
                If .NobleRep < 0 Then .NobleRep = 0
            Else
                .NobleRep = .NobleRep + vlNoble
                If .NobleRep > MAXREP Then .NobleRep = MAXREP
            End If
        End With
    End If
    
    If criminal(AttackerIndex) Then
        If UserList(AttackerIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(AttackerIndex)
        
        If Not EraCriminal Then Call RefreshCharStatus(AttackerIndex)
    ElseIf EraCriminal Then
        Call RefreshCharStatus(AttackerIndex)
    End If
    
    Call AllMascotasAtacanUser(AttackerIndex, victimIndex)
    Call AllMascotasAtacanUser(victimIndex, AttackerIndex)
    
    'Si la victima esta saliendo se cancela la salida
    Call CancelExit(victimIndex)
    Call FlushBuffer(victimIndex)
End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    'Reaccion de las mascotas
    Dim iCount As Integer
    
    For iCount = 1 To MAXMASCOTAS
        If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(victim).Name
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
        End If
    Next iCount
End Sub

Public Function PuedeAtacar(ByVal AttackerIndex As Integer, ByVal victimIndex As Integer) As Boolean
'***************************************************
'Autor: Unknown
'Last Modification: 02/04/2010
'Returns true if the AttackerIndex is allowed to attack the VictimIndex.
'24/01/2007 Pablo (ToxicWaste) - Ordeno todo y agrego situacion de Defensa en ciudad Armada y Caos.
'24/02/2009: ZaMa - Los usuarios pueden atacarse entre si.
'02/04/2010: ZaMa - Los armadas no pueden atacar nunca a los ciudas, salvo que esten atacables.
'***************************************************
On Error GoTo Errhandler

    'MUY importante el orden de estos "IF"...
    
    'Estas muerto no podes atacar
    If UserList(AttackerIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    
    'No podes atacar a alguien muerto
    If UserList(victimIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a un esp�ritu.", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    
    ' No podes atacar si estas en consulta
    If UserList(AttackerIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    
    ' No podes atacar si esta en consulta
    If UserList(victimIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estan en consulta.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If

    ' Evento 2vs2
    If Not eventAttack(AttackerIndex, victimIndex) Then
        PuedeAtacar = False
        Exit Function
    End If
    
    If UserList(AttackerIndex).flags.SlotEvent > 0 Then
        If Events(UserList(AttackerIndex).flags.SlotEvent).TimeCount > 0 Then
            WriteConsoleMsg AttackerIndex, "No puedes atacar hasta que no termine la cuenta regresiva.", FontTypeNames.FONTTYPE_INFO
            PuedeAtacar = False
            Exit Function
        End If
        
        If Events(UserList(AttackerIndex).flags.SlotEvent).Users(UserList(AttackerIndex).flags.SlotUserEvent).Team > 0 Then
            
            If Not CanAttackUserEvent(AttackerIndex, victimIndex) Then
                WriteConsoleMsg AttackerIndex, "No puedes atacar a tu compa�ero", FontTypeNames.FONTTYPE_INFO
                PuedeAtacar = False
                Exit Function
            End If
            
        End If
    End If

    'Estamos en una Arena? o un trigger zona segura?
    Select Case TriggerZonaPelea(AttackerIndex, victimIndex)
        Case eTrigger6.TRIGGER6_PERMITE
            PuedeAtacar = (UserList(victimIndex).flags.AdminInvisible = 0)
            Exit Function
        
        Case eTrigger6.TRIGGER6_PROHIBE
            PuedeAtacar = False
            Exit Function
        
        Case eTrigger6.TRIGGER6_AUSENTE
            'Si no estamos en el Trigger 6 entonces es imposible atacar un gm
            If (UserList(victimIndex).flags.Privilegios And PlayerType.User) = 0 Then
                If UserList(victimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(AttackerIndex, "El ser es demasiado poderoso.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = False
                Exit Function
            End If
    End Select
    
    'Ataca un ciudadano?
    If Not criminal(victimIndex) Then
        ' El atacante es ciuda?
        If Not criminal(AttackerIndex) Then
            ' El atacante es armada?
            If esArmada(AttackerIndex) Then
                ' La victima es armada?
                If esArmada(victimIndex) Then
                    ' No puede
                    Call WriteConsoleMsg(AttackerIndex, "Los soldados del ej�rcito real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
                    Exit Function
                End If
            End If
            
            ' Ciuda (o army) atacando a otro ciuda (o army)
            If UserList(victimIndex).flags.AtacablePor = AttackerIndex Then
                ' Se vuelve atacable.
                If ToogleToAtackable(AttackerIndex, victimIndex, False) Then
                    PuedeAtacar = True
                    Exit Function
                End If
            End If
        End If
    ' Ataca a un criminal
    Else
        'Sos un Caos atacando otro caos?
        If esCaos(victimIndex) Then
            If esCaos(AttackerIndex) Then
                Call WriteConsoleMsg(AttackerIndex, "Los miembros de la legi�n oscura tienen prohibido atacarse entre s�.", FontTypeNames.FONTTYPE_WARNING)
                Exit Function
            End If
        End If
    End If
    
    'Tenes puesto el seguro?
    If UserList(AttackerIndex).flags.Seguro Then
        If Not criminal(victimIndex) Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar ciudadanos, para hacerlo debes desactivar el seguro.", FontTypeNames.FONTTYPE_WARNING)
            PuedeAtacar = False
            Exit Function
        End If
    Else
        ' Un ciuda es atacado
        If Not criminal(victimIndex) Then
            ' Por un armada sin seguro
            If esArmada(AttackerIndex) Then
                ' No puede
                Call WriteConsoleMsg(AttackerIndex, "Los soldados del ej�rcito real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = False
                Exit Function
            End If
        End If
    End If
    
    'Estas en un Mapa Seguro?
    If MapInfo(UserList(victimIndex).Pos.map).Pk = False Then
        If esArmada(AttackerIndex) Then
            If UserList(AttackerIndex).Faccion.RecompensasReal > 11 Then
                If UserList(victimIndex).Pos.map = 58 Or UserList(victimIndex).Pos.map = 59 Or UserList(victimIndex).Pos.map = 60 Then
                Call WriteConsoleMsg(victimIndex, "�Huye de la ciudad! Est�s siendo atacado y no podr�s defenderte.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = True 'Beneficio de Armadas que atacan en su ciudad.
                Exit Function
                End If
            End If
        End If
        If esCaos(AttackerIndex) Then
            If UserList(AttackerIndex).Faccion.RecompensasCaos > 11 Then
                If UserList(victimIndex).Pos.map = 151 Or UserList(victimIndex).Pos.map = 156 Then
                Call WriteConsoleMsg(victimIndex, "�Huye de la ciudad! Est�s siendo atacado y no podr�s defenderte.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = True 'Beneficio de Caos que atacan en su ciudad.
                Exit Function
                End If
            End If
        End If
        Call WriteConsoleMsg(AttackerIndex, "Esta es una zona segura, aqu� no puedes atacar a otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    
    'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
    If MapData(UserList(victimIndex).Pos.map, UserList(victimIndex).Pos.X, UserList(victimIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Or _
        MapData(UserList(AttackerIndex).Pos.map, UserList(AttackerIndex).Pos.X, UserList(AttackerIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes pelear aqu�.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    
    PuedeAtacar = True
Exit Function

Errhandler:
    Call LogError("Error en PuedeAtacar. Error " & Err.Number & " : " & Err.Description)
End Function

Public Function PuedeAtacarNPC(ByVal AttackerIndex As Integer, ByVal NpcIndex As Integer, _
                Optional ByVal Paraliza As Boolean = False, Optional ByVal IsPet As Boolean = False) As Boolean
'***************************************************
'Autor: Unknown Author (Original version)
'Returns True if AttackerIndex can attack the NpcIndex
'Last Modification: 04/07/2010
'24/01/2007 Pablo (ToxicWaste) - Orden y correcci�n de ataque sobre una mascota y guardias
'14/08/2007 Pablo (ToxicWaste) - Reescribo y agrego TODOS los casos posibles cosa de usar
'esta funci�n para todo lo referente a ataque a un NPC. Ya sea Magia, F�sico o a Distancia.
'16/11/2009: ZaMa - Agrego validacion de pertenencia de npc.
'02/04/2010: ZaMa - Los armadas ya no peuden atacar npcs no hotiles.
'23/05/2010: ZaMa - El inmo/para renuevan el timer de pertenencia si el ataque fue a un npc propio.
'04/07/2010: ZaMa - Ahora no se puede apropiar del dragon de dd.
'***************************************************
 
On Error GoTo Errhandler
 
    With Npclist(NpcIndex)
   
        'Estas muerto?
        If UserList(AttackerIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(AttackerIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
       
        'Sos consejero?
        If UserList(AttackerIndex).flags.Privilegios And PlayerType.Consejero Then
            'No pueden atacar NPC los Consejeros.
            Exit Function
        End If
       
      'Estas en modo Combate?
    If Not UserList(AttackerIndex).flags.ModoCombate Then
        Call WriteConsoleMsg(AttackerIndex, "Para realizar esta accion debes activar el modo combate, puedes hacerlo con la tecla ""C""", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacarNPC = False
        Exit Function
    End If
    
       
        ' No podes atacar si estas en consulta
        If UserList(AttackerIndex).flags.EnConsulta Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar npcs mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
       
        'Es una criatura atacable?
        If .Attackable = 0 Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar esta criatura.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
       
        'Es valida la distancia a la cual estamos atacando?
        If Distancia(UserList(AttackerIndex).Pos, .Pos) >= MAXDISTANCIAARCO Then
           Call WriteConsoleMsg(AttackerIndex, "Est�s muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
           Exit Function
        End If
       
        'Es una criatura No-Hostil?
        If .Hostile = 0 Then
            'Es Guardia del Caos?
            If .NPCtype = eNPCType.Guardiascaos Then
                'Lo quiere atacar un caos?
                If esCaos(AttackerIndex) Then
                    Call WriteConsoleMsg(AttackerIndex, "No puedes atacar Guardias del Caos siendo de la legi�n oscura.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            'Es guardia Real?
            ElseIf .NPCtype = eNPCType.GuardiaReal Then
                'Lo quiere atacar un Armada?
                If esArmada(AttackerIndex) Then
                    Call WriteConsoleMsg(AttackerIndex, "No puedes atacar Guardias Reales siendo del ej�rcito real.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
                'Tienes el seguro puesto?
                If UserList(AttackerIndex).flags.Seguro Then
                    Call WriteConsoleMsg(AttackerIndex, "Para poder atacar Guardias Reales debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                Else
                    Call WriteConsoleMsg(AttackerIndex, "�Atacaste un Guardia Real! Eres un criminal.", FontTypeNames.FONTTYPE_INFO)
                    Call VolverCriminal(AttackerIndex)
                    PuedeAtacarNPC = True
                    Exit Function
                End If
       
            'No era un Guardia, asi que es una criatura No-Hostil com�n.
            'Para asegurarnos que no sea una Mascota:
            ElseIf .MaestroUser = 0 Then
                'Si sos ciudadano tenes que quitar el seguro para atacarla.
                If Not criminal(AttackerIndex) Then
                   
                    ' Si sos armada no podes atacarlo directamente
                    If esArmada(AttackerIndex) Then
                        Call WriteConsoleMsg(AttackerIndex, "Los miembros del ej�rcito real no pueden atacar npcs no hostiles.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    End If
               
                    'Sos ciudadano, tenes el seguro puesto?
                    If UserList(AttackerIndex).flags.Seguro Then
                        Call WriteConsoleMsg(AttackerIndex, "Para atacar a este NPC debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    Else
                        'No tiene seguro puesto. Puede atacar pero es penalizado.
                        Call WriteConsoleMsg(AttackerIndex, "Atacaste un NPC no-hostil. Contin�a haci�ndolo y te podr�s convertir en criminal.", FontTypeNames.FONTTYPE_INFO)
                        'NicoNZ: Cambio para que al atacar npcs no hostiles no bajen puntos de nobleza
                        Call DisNobAuBan(AttackerIndex, 0, 1000)
                        PuedeAtacarNPC = True
                        Exit Function
                    End If
                End If
            End If
        End If
   
   
        Dim MasterIndex As Integer
        MasterIndex = .MaestroUser
       
        'Es el NPC mascota de alguien?
        If MasterIndex > 0 Then
           
            ' Due�o de la mascota ciuda?
            If Not criminal(MasterIndex) Then
               
                ' Atacante ciuda?
                If Not criminal(AttackerIndex) Then
                   
                    ' Si esta en estado atacable puede atacar su mascota sin problemas
                    If UserList(MasterIndex).flags.AtacablePor = AttackerIndex Then
                        ' Toogle to atacable and restart the timer
                        Call ToogleToAtackable(AttackerIndex, MasterIndex)
                        PuedeAtacarNPC = True
                        Exit Function
                    End If
                   
                    'Atacante armada?
                    If esArmada(AttackerIndex) Then
                        'El atacante es Armada y esta intentando atacar mascota de un Ciudadano
                        Call WriteConsoleMsg(AttackerIndex, "Los miembros del ej�rcito real no pueden atacar mascotas de ciudadanos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    End If
                   
                    'El atacante es Ciudadano y esta intentando atacar mascota de un Ciudadano.
                    If UserList(AttackerIndex).flags.Seguro Then
                        'El atacante tiene el seguro puesto. No puede atacar.
                        Call WriteConsoleMsg(AttackerIndex, "Para atacar mascotas de ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    Else
                        'El atacante no tiene el seguro puesto. Recibe penalizaci�n.
                        Call WriteConsoleMsg(AttackerIndex, "Has atacado la Mascota de un ciudadano. Eres un criminal.", FontTypeNames.FONTTYPE_INFO)
                        Call VolverCriminal(AttackerIndex)
                        PuedeAtacarNPC = True
                        Exit Function
                    End If
                Else
                    'El atacante es criminal y quiere atacar un elemental ciuda, pero tiene el seguro puesto (NicoNZ)
                    If UserList(AttackerIndex).flags.Seguro Then
                        Call WriteConsoleMsg(AttackerIndex, "Para atacar mascotas de ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    End If
                End If
           
            'Es mascota de un caos?
            ElseIf esCaos(MasterIndex) Then
                'Es Caos el Due�o.
                If esCaos(AttackerIndex) Then
                    'Un Caos intenta atacar una criatura de un Caos. No puede atacar.
                    Call WriteConsoleMsg(AttackerIndex, "Los miembros de la legi�n oscura no pueden atacar mascotas de otros legionarios. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            End If
           
        ' No es mascota de nadie, le pertenece a alguien?
        ElseIf .Owner > 0 Then
       
            Dim OwnerUserIndex As Integer
            OwnerUserIndex = .Owner
           
            ' Puede atacar a su propia criatura!
            If OwnerUserIndex = AttackerIndex Then
                PuedeAtacarNPC = True
                Call IntervaloPerdioNpc(OwnerUserIndex, True) ' Renuevo el timer
                Exit Function
            End If
           
            ' Esta compartiendo el npc con el atacante? => Puede atacar!
            If UserList(OwnerUserIndex).flags.ShareNpcWith = AttackerIndex Then
                PuedeAtacarNPC = True
                Exit Function
            End If
           
            ' Si son del mismo clan o party, pueden atacar (No renueva el timer)
            If Not SameClan(OwnerUserIndex, AttackerIndex) And Not SameParty(OwnerUserIndex, AttackerIndex) Then
           
                ' Si se le agoto el tiempo
                If IntervaloPerdioNpc(OwnerUserIndex) Then ' Se lo roba :P
                    Call PerdioNpc(OwnerUserIndex)
                    Call ApropioNpc(AttackerIndex, NpcIndex)
                    PuedeAtacarNPC = True
                    Exit Function
                   
                ' Si lanzo un hechizo de para o inmo
                ElseIf Paraliza Then
               
                    ' Si ya esta paralizado o inmobilizado, no puedo inmobilizarlo de nuevo
                    If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then
                       
                        'TODO_ZAMA: Si dejo esto asi, los pks con seguro peusto van a poder inmobilizar criaturas con due�o
                        ' Si es pk neutral, puede hacer lo que quiera :P.
                        If Not criminal(AttackerIndex) And Not criminal(OwnerUserIndex) Then
                       
                             'El atacante es Armada
                            If esArmada(AttackerIndex) Then
                               
                                 'Intententa paralizar un npc de un armada?
                                If esArmada(OwnerUserIndex) Then
                                    'El atacante es Armada y esta intentando paralizar un npc de un armada: No puede
                                    Call WriteConsoleMsg(AttackerIndex, "Los miembros del Ej�rcito Real no pueden paralizar criaturas ya paralizadas pertenecientes a otros miembros del Ej�rcito Real", FontTypeNames.FONTTYPE_INFO)
                                    Exit Function
                               
                                'El atacante es Armada y esta intentando paralizar un npc de un ciuda
                                Else
                                    ' Si tiene seguro no puede
                                    If UserList(AttackerIndex).flags.Seguro Then
                                        Call WriteConsoleMsg(AttackerIndex, "Para paralizar criaturas ya paralizadas pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Function
                                    Else
                                        ' Si ya estaba atacable, no podr� atacar a un npc perteneciente a otro ciuda
                                        If ToogleToAtackable(AttackerIndex, OwnerUserIndex) Then
                                            Call WriteConsoleMsg(AttackerIndex, "Has paralizado la criatura de un ciudadano, ahora eres atacable por �l.", FontTypeNames.FONTTYPE_INFO)
                                            PuedeAtacarNPC = True
                                        End If
                                       
                                        Exit Function
                                       
                                    End If
                                End If
                               
                            ' El atacante es ciuda
                            Else
                                'El atacante tiene el seguro puesto, no puede paralizar
                                If UserList(AttackerIndex).flags.Seguro Then
                                    Call WriteConsoleMsg(AttackerIndex, "Para paralizar criaturas ya paralizadas pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Function
                                   
                                'El atacante no tiene el seguro puesto, ataca.
                                Else
                                    ' Si ya estaba atacable, no podr� atacar a un npc perteneciente a otro ciuda
                                    If ToogleToAtackable(AttackerIndex, OwnerUserIndex) Then
                                        Call WriteConsoleMsg(AttackerIndex, "Has paralizado la criatura de un ciudadano, ahora eres atacable por �l.", FontTypeNames.FONTTYPE_INFO)
                                        PuedeAtacarNPC = True
                                    End If
                                   
                                    Exit Function
                                End If
                            End If
                           
                        ' Al menos uno de los dos es criminal
                        Else
                            ' Si ambos son caos
                            If esCaos(AttackerIndex) And esCaos(OwnerUserIndex) Then
                                'El atacante es Caos y esta intentando paralizar un npc de un Caos
                                Call WriteConsoleMsg(AttackerIndex, "Los miembros de la legi�n oscura no pueden paralizar criaturas ya paralizadas por otros legionarios.", FontTypeNames.FONTTYPE_INFO)
                                Exit Function
                            End If
                        End If
                   
                    ' El npc no esta inmobilizado ni paralizado
                    Else
                        ' Si no tiene due�o, puede apropiarselo
                        If OwnerUserIndex = 0 Then
                       
                            ' Siempre que no posea uno ya (el inmo/para no cambia pertenencia de npcs).
                            If UserList(AttackerIndex).flags.OwnedNpc = 0 Then
                                Call ApropioNpc(AttackerIndex, NpcIndex)
                            End If
                           
                        ' Si inmobiliza a su propio npc, renueva el timer
                        ElseIf OwnerUserIndex = AttackerIndex Then
                            Call IntervaloPerdioNpc(OwnerUserIndex, True) ' Renuevo el timer
                        End If
                       
                        ' Siempre se pueden paralizar/inmobilizar npcs con o sin due�o
                        ' que no tengan ese estado
                        PuedeAtacarNPC = True
                        Exit Function
 
                    End If
                   
                ' No lanz� hechizos inmobilizantes
                Else
                   
                    ' El npc le pertenece a un ciudadano
                    If Not criminal(OwnerUserIndex) Then
                       
                        'El atacante es Armada y esta intentando atacar un npc de un Ciudadano
                        If esArmada(AttackerIndex) Then
                        
                            If Not .flags.TeamEvent > 0 Then
                                'Intententa atacar un npc de un armada?
                                If esArmada(OwnerUserIndex) Then
                                    'El atacante es Armada y esta intentando atacar el npc de un armada: No puede
                                    Call WriteConsoleMsg(AttackerIndex, "Los miembros del Ej�rcito Real no pueden atacar criaturas pertenecientes a otros miembros del Ej�rcito Real", FontTypeNames.FONTTYPE_INFO)
                                    Exit Function
                               
                                'El atacante es Armada y esta intentando atacar un npc de un ciuda
                                Else
                                   
                                    ' Si tiene seguro no puede
                                    If UserList(AttackerIndex).flags.Seguro Then
                                        Call WriteConsoleMsg(AttackerIndex, "Para atacar criaturas ya pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Function
                                    Else
                            
                                       
                                        Exit Function
                                    End If
                                End If
                            End If
                           
                        ' No es aramda, puede ser criminal o ciuda
                        Else
                           
                            'El atacante es Ciudadano y esta intentando atacar un npc de un Ciudadano.
                            If Not criminal(AttackerIndex) Then
                               
                                If UserList(AttackerIndex).flags.Seguro Then
                                    'El atacante tiene el seguro puesto. No puede atacar.
                                    Call WriteConsoleMsg(AttackerIndex, "Para atacar criaturas pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Function
                               
                                'El atacante no tiene el seguro puesto, ataca.
                                Else
                                    If ToogleToAtackable(AttackerIndex, OwnerUserIndex) Then
                                        Call WriteConsoleMsg(AttackerIndex, "Has atacado a la criatura de un ciudadano, ahora eres atacable por �l.", FontTypeNames.FONTTYPE_INFO)
                                        PuedeAtacarNPC = True
                                    End If
                                   
                                    Exit Function
                                End If
                               
                            'El atacante es criminal y esta intentando atacar un npc de un Ciudadano.
                            Else
                                ' Es criminal atacando un npc de un ciuda, con seguro puesto.
                                If UserList(AttackerIndex).flags.Seguro Then
                                    Call WriteConsoleMsg(AttackerIndex, "Para atacar criaturas pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Function
                                End If
                               
                                PuedeAtacarNPC = True
                            End If
                        End If
                       
                    ' Es npc de un criminal
                    Else
                        If Not .flags.TeamEvent > 0 Then
                            If esCaos(OwnerUserIndex) Then
                                'Es Caos el Due�o.
                                If esCaos(AttackerIndex) Then
                                    'Un Caos intenta atacar una npc de un Caos. No puede atacar.
                                    Call WriteConsoleMsg(AttackerIndex, "Los miembros de la Legi�n Oscura no pueden atacar criaturas de otros legionarios. ", FontTypeNames.FONTTYPE_INFO)
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End If
           
        ' Si no tiene due�o el npc, se lo apropia
        Else
            ' Solo pueden apropiarse de npcs los caos, armadas o ciudas.
            If Not criminal(AttackerIndex) Or esCaos(AttackerIndex) Then
                ' No puede apropiarse de los pretos!
                If Npclist(NpcIndex).NPCtype <> eNPCType.pretoriano Then
                    ' No puede apropiarse del dragon de dd!
                    If Npclist(NpcIndex).NPCtype <> Dragon Then
                        ' Si es una mascota atacando, no se apropia del npc
                        If Not IsPet Then
                            ' No es due�o de ningun npc => Se lo apropia.
                            If UserList(AttackerIndex).flags.OwnedNpc = 0 Then
                                Call ApropioNpc(AttackerIndex, NpcIndex)
                            ' Es due�o de un npc, pero no puede ser de este porque no tiene propietario.
                            Else
                                ' Se va a adue�ar del npc (y perder el otro) solo si no inmobiliza/paraliza
                                If Not Paraliza Then Call ApropioNpc(AttackerIndex, NpcIndex)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        
        If (UserList(AttackerIndex).flags.SlotEvent) > 0 And (.flags.TeamEvent > 0) Then
            If Events(UserList(AttackerIndex).flags.SlotEvent).Modality = CastleMode Then
                If Not EventosDS.CanAttackReyCastle(AttackerIndex, NpcIndex) Then
                    WriteConsoleMsg AttackerIndex, "No puedes atacar a tu rey", FontTypeNames.FONTTYPE_FIGHT
                    Exit Function
                End If
            
            End If
        
        End If
        
    End With
   
    'Es el Rey Preatoriano?
    If esPretoriano(NpcIndex) = 4 Then
        If pretorianosVivos > 0 Then
            Call WriteConsoleMsg(AttackerIndex, "Debes matar al resto del ej�rcito antes de atacar al rey.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Function
        End If
    End If
   
    PuedeAtacarNPC = True
       
    Exit Function
       
Errhandler:
   
    Dim AtckName As String
    Dim OwnerName As String
 
    If AttackerIndex > 0 Then AtckName = UserList(AttackerIndex).Name
    If OwnerUserIndex > 0 Then OwnerName = UserList(OwnerUserIndex).Name
   
    Call LogError("Error en PuedeAtacarNpc. Erorr: " & Err.Number & " - " & Err.Description & " Atacante: " & _
                   AttackerIndex & "-> " & AtckName & ". Owner: " & OwnerUserIndex & "-> " & OwnerName & _
                   ". NpcIndex: " & NpcIndex & ".")
End Function

Private Function SameClan(ByVal UserIndex As Integer, ByVal OtherUserIndex As Integer) As Boolean
'***************************************************
'Autor: ZaMa
'Returns True if both players belong to the same clan.
'Last Modification: 16/11/2009
'***************************************************
    SameClan = (UserList(UserIndex).GuildIndex = UserList(OtherUserIndex).GuildIndex) And _
                UserList(UserIndex).GuildIndex <> 0
End Function

Private Function SameParty(ByVal UserIndex As Integer, ByVal OtherUserIndex As Integer) As Boolean
'***************************************************
'Autor: ZaMa
'Returns True if both players belong to the same party.
'Last Modification: 16/11/2009
'***************************************************
    SameParty = UserList(UserIndex).PartyIndex = UserList(OtherUserIndex).PartyIndex And _
                UserList(UserIndex).PartyIndex <> 0
End Function

Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDa�o As Long)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/09/06 Nacho
'Reescribi gran parte del Sub
'Ahora, da toda la experiencia del npc mientras este vivo.
'***************************************************
    Dim ExpaDar As Long
    
    '[Nacho] Chekeamos que las variables sean validas para las operaciones
    If ElDa�o <= 0 Then ElDa�o = 0
    If Npclist(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub
    If ElDa�o > Npclist(NpcIndex).Stats.MinHp Then ElDa�o = Npclist(NpcIndex).Stats.MinHp
    
    
    'Npclist(NpcIndex).Stats.MinHp = 1
    If ElDa�o < 0 Then ElDa�o = Npclist(NpcIndex).Stats.MinHp
    '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
    ExpaDar = CLng(ElDa�o * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHp))
    If ExpaDar <= 0 Then Exit Sub
    
    '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
            'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
            'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
    If ExpaDar > Npclist(NpcIndex).flags.ExpCount Then
        ExpaDar = Npclist(NpcIndex).flags.ExpCount
        Npclist(NpcIndex).flags.ExpCount = 0
    Else
        Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).flags.ExpCount - ExpaDar
    End If
    
    '[Nacho] Le damos la exp al user
    If ExpaDar > 0 Then
    
        If UserList(UserIndex).PartyIndex > 0 Then
            Call mdParty.ObtenerExito(UserIndex, ExpaDar, Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y)
        Else
            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar
            If UserList(UserIndex).Stats.Exp > MAXEXP Then _
                UserList(UserIndex).Stats.Exp = MAXEXP
            Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpaDar & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
            If UserList(UserIndex).flags.Oro = 1 Then
            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + (ExpaDar * 0.35)
                 WriteConsoleMsg UserIndex, "Aumento de exp 35% > Has ganado " & (ExpaDar * 0.35) & " puntos de experiencia.", FontTypeNames.fonttype_dios
        End If
        End If
        
        Call CheckUserLevel(UserIndex)
    End If
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'TODO: Pero que rebuscado!!
'Nigo:  Te lo redise�e, pero no te borro el TODO para que lo revises.
On Error GoTo Errhandler
    Dim tOrg As eTrigger
    Dim tDst As eTrigger
    
    tOrg = MapData(UserList(Origen).Pos.map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).trigger
    tDst = MapData(UserList(Destino).Pos.map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).trigger
    
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If

Exit Function
Errhandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & Err.Description)
End Function

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim ObjInd As Integer
    
    ObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    If ObjInd > 0 Then
        If ObjData(ObjInd).proyectil = 1 Then
            ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
        End If
        
        If ObjInd > 0 Then
            If ObjData(ObjInd).Envenena = 1 Then
                
                If RandomNumber(1, 100) < 60 Then
                    UserList(VictimaIndex).flags.Envenenado = 1
                    Call WriteConsoleMsg(VictimaIndex, "��" & UserList(AtacanteIndex).Name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(AtacanteIndex, "��Has envenenado a " & UserList(VictimaIndex).Name & "!!", FontTypeNames.FONTTYPE_FIGHT)
                End If
            End If
        End If
    End If
    
    Call FlushBuffer(VictimaIndex)
End Sub
