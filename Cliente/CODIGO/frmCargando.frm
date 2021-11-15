VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7980
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCargando.frx":0000
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   532
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   480
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Status 
      Height          =   1545
      Left            =   1440
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   8520
      Visible         =   0   'False
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   2725
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmCargando.frx":1C511
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox LOGO 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   960
      Picture         =   "frmCargando.frx":1C58F
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   9480
      Width           =   9600
   End
   Begin VB.Image barra 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   1320
      Picture         =   "frmCargando.frx":37419
      Top             =   1890
      Width           =   5415
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Desterium  AO 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software;you can redistribute it and/or modify
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
'La Plata - Pcia, Buenos Aiçres - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit
Dim F As Integer


Private Sub Form_Load()
barra.Width = 1
End Sub



Private Sub LOGO_KeyPress(KeyAscii As Integer)
Debug.Print 2
End Sub

Private Sub Status_KeyPress(KeyAscii As Integer)
Debug.Print 1
End Sub
Function Analizar()
            On Error Resume Next
           
            Dim iX As Integer
            Dim tX As Integer
            Dim DifX As Integer
           
'LINK1            'Variable que contiene el numero de actualización correcto del servidor
    iX = Inet1.OpenURL("https://Desterium -ao.ucoz.net/VEREXE.txt") 'Host
    tX = LeerInt(App.path & "\INIT\Update.ini")
    DifX = iX - tX
 
            If Not (DifX = 0) Then
If MsgBox("¡SE ABRIRÁ EL ACTUALIZADOR, NECESITAS ACTUALIZAR EL CLIENTE!.") Then
Call ShellExecute(Me.hwnd, "open", App.path & "/Actualizador.exe", "", "", 1)
End
End If
End If

End Function
Private Function LeerInt(ByVal Ruta As String) As Integer
    F = FreeFile
    Open Ruta For Input As F
    LeerInt = Input$(LOF(F), #F)
    Close #F
End Function
 
Private Sub GuardarInt(ByVal Ruta As String, ByVal data As Integer)
    F = FreeFile
    Open Ruta For Output As F
    Print #F, data
    Close #F
End Sub
