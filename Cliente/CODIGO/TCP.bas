Attribute VB_Name = "Mod_TCP"
'Desterium  AO 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Desterium  AO is based on Baronsoft's VB6 Online RPG
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
Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean
Public LlegoFama As Boolean



Public Function PuedoQuitarFoco() As Boolean
PuedoQuitarFoco = True
'PuedoQuitarFoco = Not frmEstadisticas.Visible And _
'                 Not frmGuildAdm.Visible And _
'                 Not frmGuildDetails.Visible And _
'                 Not frmGuildBrief.Visible And _
'                 Not frmGuildFoundation.Visible And _
'                 Not frmGuildLeader.Visible And _
'                 Not frmCharInfo.Visible And _
'                 Not frmGuildNews.Visible And _
'                 Not frmGuildSol.Visible And _
'                 Not frmCommet.Visible And _
'                 Not frmPeaceProp.Visible
'
End Function

Sub Login()
    If EstadoLogin = E_MODO.Normal Then
        Call WriteConectarUsuarioE
    ElseIf EstadoLogin = E_MODO.CrearNuevoPj Then
        Call WriteLogeaNuevoPj
    ElseIf EstadoLogin = E_MODO.BorrarPJ Then
        Call WriteKillChar
    ElseIf EstadoLogin = E_MODO.RecuperarPJ Then
        Call WriteRenewPassChar
    End If
   
    DoEvents
   
    'Call FlushBuffer
    
    Dim sndData As String
    
    With outgoingData
        If .Length = 0 Then _
            Exit Sub
        
        sndData = .ReadASCIIStringFixed(.Length) 'Leo el paquete
        Randomize
        CRC = RandomNumber(0, 450)
        sndData = ConvertirFlush(sndData)        'Lo encripto
        Call WritePaqueteEncriptado              'Ingreso las bases para el server
        sndData = .ReadASCIIStringFixed(.Length) & sndData  'Agrego las bases al paquete
        .WriteASCIIStringFixed (sndData)         'Lo escribo
        Call FlushBuffer                         'Lo mando
    End With
End Sub
