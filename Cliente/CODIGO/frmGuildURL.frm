VERSION 5.00
Begin VB.Form frmGuildURL 
   BorderStyle     =   0  'None
   Caption         =   "Oficial Web Site"
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6225
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
   Picture         =   "frmGuildURL.frx":0000
   ScaleHeight     =   95
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUrl 
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
      Height          =   345
      Left            =   240
      TabIndex        =   0
      Top             =   525
      Width           =   5745
   End
   Begin VB.Image imgAceptar 
      Height          =   375
      Left            =   2640
      Tag             =   "1"
      Top             =   960
      Width           =   1320
   End
End
Attribute VB_Name = "frmGuildURL"
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

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
   ' Me.Picture = LoadPicture(App.path & "\Recursos\VentanaUrlClan.jpg")
    
End Sub
Private Sub imgAceptar_Click()
    If txtUrl.Text <> "" Then _
        Call WriteGuildNewWebsite(txtUrl.Text)
    
    Unload Me
End Sub
