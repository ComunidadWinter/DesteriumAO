VERSION 5.00
Begin VB.Form frmGuildBrief 
   BorderStyle     =   0  'None
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   7620
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
   Picture         =   "frmGuildBrief.frx":0000
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   508
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "&H8000000A&"
   Begin VB.TextBox Desc 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   315
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   6045
      Width           =   6915
   End
   Begin VB.Image ImgOfrecerPaz 
      Height          =   375
      Left            =   1200
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image ImgOfrecerAlianza 
      Height          =   375
      Left            =   2520
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Image ImgDeclararGuerra 
      Height          =   375
      Left            =   4320
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Image imgSolicitarIngreso 
      Height          =   375
      Left            =   6000
      Tag             =   "1"
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Image imgCerrar 
      Height          =   360
      Left            =   0
      Tag             =   "1"
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   390
      TabIndex        =   18
      Top             =   3840
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   390
      TabIndex        =   17
      Top             =   4080
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   390
      TabIndex        =   16
      Top             =   4320
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   390
      TabIndex        =   15
      Top             =   4560
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   390
      TabIndex        =   14
      Top             =   4800
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   390
      TabIndex        =   13
      Top             =   5040
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   390
      TabIndex        =   12
      Top             =   5280
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   390
      TabIndex        =   11
      Top             =   5520
      Width           =   6735
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   555
      Width           =   4695
   End
   Begin VB.Label fundador 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1410
      TabIndex        =   9
      Top             =   1185
      Width           =   2775
   End
   Begin VB.Label creacion 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   870
      Width           =   1455
   End
   Begin VB.Label lider 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5430
      TabIndex        =   7
      Top             =   870
      Width           =   3135
   End
   Begin VB.Label web 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1260
      TabIndex        =   6
      Top             =   1485
      Width           =   2655
   End
   Begin VB.Label Miembros 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   2445
      Width           =   1935
   End
   Begin VB.Label eleccion 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6075
      TabIndex        =   4
      Top             =   555
      Width           =   1815
   End
   Begin VB.Label lblAlineacion 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1485
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Enemigos 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1890
      TabIndex        =   2
      Top             =   1815
      Width           =   2175
   End
   Begin VB.Label Aliados 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1785
      TabIndex        =   1
      Top             =   2115
      Width           =   1575
   End
   Begin VB.Label antifaccion 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   4320
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "frmGuildBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M?rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat?as Fernando Peque?o
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

Private clsFormulario As clsFormMovementManager

Private cBotonGuerra As clsGraphicalButton
Private cBotonAlianza As clsGraphicalButton
Private cBotonPaz As clsGraphicalButton
Private cBotonSolicitarIngreso As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Public EsLeader As Boolean

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
  '  Me.Picture = LoadPicture(App.path & "\Recursos\VentanaDetallesClan.jpg")
    

    
End Sub


Private Sub ImgOfrecerAlianza_Click()
    frmCommet.Nombre = Nombre.Caption
    frmCommet.T = TIPO.ALIANZA
    Call frmCommet.Show(vbModal, frmGuildBrief)
End Sub

Private Sub ImgOfrecerpaz_Click()
    frmCommet.Nombre = Nombre.Caption
    frmCommet.T = TIPO.PAZ
    Call frmCommet.Show(vbModal, frmGuildBrief)
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgDeclararGuerra_Click()
    Call WriteGuildDeclareWar(Nombre.Caption)
    Unload Me
End Sub



Private Sub imgSolicitarIngreso_Click()
    Call frmGuildSol.RecieveSolicitud(Nombre.Caption)
    Call frmGuildSol.Show(vbModal, frmGuildBrief)
End Sub

Private Sub Label1_Click()

End Sub

