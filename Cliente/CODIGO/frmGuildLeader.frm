VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   0  'None
   Caption         =   "Administración del Clan"
   ClientHeight    =   7425
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   5970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGuildLeader.frx":0000
   ScaleHeight     =   495
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   398
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFiltrarMiembros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3000
      TabIndex        =   6
      Top             =   2310
      Width           =   2685
   End
   Begin VB.TextBox txtFiltrarClanes 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   135
      TabIndex        =   5
      Top             =   2310
      Width           =   2685
   End
   Begin VB.TextBox txtguildnews 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3390
      Width           =   5475
   End
   Begin VB.ListBox solicitudes 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      ItemData        =   "frmGuildLeader.frx":22714
      Left            =   150
      List            =   "frmGuildLeader.frx":22716
      TabIndex        =   2
      ToolTipText     =   "1"
      Top             =   5280
      Width           =   2715
   End
   Begin VB.ListBox members 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      ItemData        =   "frmGuildLeader.frx":22718
      Left            =   3045
      List            =   "frmGuildLeader.frx":2271A
      TabIndex        =   1
      Top             =   540
      Width           =   2595
   End
   Begin VB.ListBox guildslist 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      ItemData        =   "frmGuildLeader.frx":2271C
      Left            =   195
      List            =   "frmGuildLeader.frx":2271E
      TabIndex        =   0
      Top             =   540
      Width           =   2595
   End
   Begin VB.Image ImgElecciones 
      Height          =   375
      Left            =   3600
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Miembros 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1965
      TabIndex        =   4
      Top             =   6765
      Width           =   2535
   End
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   3720
      Tag             =   "1"
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Image imgPropuestasPaz 
      Height          =   375
      Left            =   3480
      Tag             =   "1"
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Image imgEditarURL 
      Height          =   495
      Left            =   3480
      Tag             =   "1"
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Image imgEditarCodex 
      Height          =   375
      Left            =   3360
      Tag             =   "1"
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Image imgActualizar 
      Height          =   390
      Left            =   2400
      Tag             =   "1"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Image imgDetallesSolicitudes 
      Height          =   375
      Left            =   120
      Tag             =   "1"
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Image imgDetallesMiembros 
      Height          =   495
      Left            =   3840
      Tag             =   "1"
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Image imgDetallesClan 
      Height          =   495
      Left            =   120
      Tag             =   "1"
      Top             =   2640
      Width           =   1935
   End
End
Attribute VB_Name = "frmGuildLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
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

Private Const MAX_NEWS_LENGTH As Integer = 512
Private clsFormulario As clsFormMovementManager

Private cBotonElecciones As clsGraphicalButton
Private cBotonActualizar As clsGraphicalButton
Private cBotonDetallesClan As clsGraphicalButton
Private cBotonDetallesMiembros As clsGraphicalButton
Private cBotonDetallesSolicitudes As clsGraphicalButton
Private cBotonEditarCodex As clsGraphicalButton
Private cBotonEditarURL As clsGraphicalButton
Private cBotonPropuestasPaz As clsGraphicalButton
Private cBotonPropuestasAlianzas As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
'    Me.Picture = LoadPicture(App.path & "\Recursos\VentanaAdministrarClan.jpg")
    
 
End Sub


Private Sub imgActualizar_Click()
    Dim k As String

    k = Replace(txtguildnews, vbCrLf, "º")
    
    Call WriteGuildUpdateNews(k)
End Sub

Private Sub imgCerrar_Click()
    Unload Me
    frmMain.SetFocus
End Sub

Private Sub imgDetallesClan_Click()
    frmGuildBrief.EsLeader = True
    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))
End Sub

Private Sub imgDetallesMiembros_Click()
    If members.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembers
    Call WriteGuildMemberInfo(members.List(members.ListIndex))
End Sub

Private Sub imgDetallesSolicitudes_Click()
    If solicitudes.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembershipRequests
    Call WriteGuildMemberInfo(solicitudes.List(solicitudes.ListIndex))
End Sub

Private Sub imgEditarCodex_Click()
    Call frmGuildDetails.Show(vbModal, frmGuildLeader)
End Sub

Private Sub imgEditarURL_Click()
    Call frmGuildURL.Show(vbModeless, frmGuildLeader)
End Sub

Private Sub imgElecciones_Click()
            If MsgBox("¿Estás seguro de abrir elecciones?", vbYesNo, "Opciones del clan") = vbYes Then _
    Call WriteGuildOpenElections
    Unload Me
End Sub


Private Sub ImgEleccion_Click()

End Sub

Private Sub imgPropuestasPaz_Click()
    Call WriteGuildPeacePropList
End Sub


Private Sub txtFiltrarClanes_Change()
Call FiltrarListaClanes(txtFiltrarClanes.Text)
End Sub

Private Sub txtFiltrarMiembros_Change()
 Call FiltrarListaMiembros(txtFiltrarMiembros.Text)
End Sub

Private Sub txtguildnews_Change()
    If Len(txtguildnews.Text) > MAX_NEWS_LENGTH Then _
        txtguildnews.Text = Left$(txtguildnews.Text, MAX_NEWS_LENGTH)
End Sub

Private Sub txtFiltrarClanes_GotFocus()
    With txtFiltrarClanes
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub FiltrarListaClanes(ByRef sCompare As String)

    Dim lIndex As Long
    
    With guildslist
        'Limpio la lista
        .Clear
        
        .Visible = False
        
        ' Recorro los arrays
        For lIndex = 0 To UBound(GuildNames)
            ' Si coincide con los patrones
            If InStr(1, UCase$(GuildNames(lIndex)), UCase$(sCompare)) Then
                ' Lo agrego a la lista
                .AddItem GuildNames(lIndex)
            End If
        Next lIndex
        
        .Visible = True
    End With

End Sub

Private Sub txtFiltrarMiembros_GotFocus()
    With txtFiltrarMiembros
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub FiltrarListaMiembros(ByRef sCompare As String)

    Dim lIndex As Long
    
    With members
        'Limpio la lista
        .Clear
        
        .Visible = False
        
        ' Recorro los arrays
        For lIndex = 0 To UBound(GuildMembers)
            ' Si coincide con los patrones
            If InStr(1, UCase$(GuildMembers(lIndex)), UCase$(sCompare)) Then
                ' Lo agrego a la lista
                .AddItem GuildMembers(lIndex)
            End If
        Next lIndex
        
        .Visible = True
    End With
End Sub



