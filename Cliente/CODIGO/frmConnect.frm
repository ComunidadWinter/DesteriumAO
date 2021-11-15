VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Estelar AO"
   ClientHeight    =   8985
   ClientLeft      =   2925
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   11400
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox checkrecu 
      BackColor       =   &H00000000&
      Caption         =   "Recordar"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   8055
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   6
      Top             =   4980
      Width           =   1035
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   4905
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4650
      Width           =   4185
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4905
      TabIndex        =   0
      Top             =   3840
      Width           =   4185
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4890
      TabIndex        =   2
      Text            =   "7666"
      Top             =   -4800
      Width           =   825
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5760
      TabIndex        =   4
      Text            =   "localhost"
      Top             =   -4800
      Width           =   1575
   End
   Begin SHDocVwCtl.WebBrowser WebAuxiliar 
      Height          =   360
      Left            =   960
      TabIndex        =   5
      Top             =   -840
      Visible         =   0   'False
      Width           =   330
      ExtentX         =   582
      ExtentY         =   635
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   5400
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   720
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   600
      Top             =   5640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   600
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label OnOff 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Comprobando..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   9240
      TabIndex        =   10
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Image Borrar 
      Height          =   495
      Left            =   600
      Top             =   6480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Recuperar 
      Height          =   495
      Left            =   600
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "crear PEJOTAAAAAAAAAAAAAAAA"
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   -7560
      Width           =   4215
   End
   Begin VB.Label imgSalir 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Label imgCrearPj 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Image imgConectarse 
      Height          =   375
      Left            =   7080
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Image imgServArgentina 
      Height          =   795
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   9240
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   -240
      Width           =   555
   End
End
Attribute VB_Name = "frmConnect"
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
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Matías Fernando Pequeño
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Código Postal 1405

Option Explicit

Private cBotonCrearPj As clsGraphicalButton
Private cBotonRecuperarPass As clsGraphicalButton
Private cBotonManual As clsGraphicalButton
Private cBotonReglamento As clsGraphicalButton
Private cBotonCodigoFuente As clsGraphicalButton
Private cBotonBorrarPj As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton
Private cBotonLeerMas As clsGraphicalButton
Private cBotonForo As clsGraphicalButton
Private cBotonConectarse As clsGraphicalButton
Private cBotonTeclas As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_Activate()
'On Error Resume Next


If ServersRecibidos Then
    If CurServer <> 0 Then
        IPTxt = ServersLst(1).Ip
        PortTxt = ServersLst(1).Puerto
    Else
        IPTxt = IPdelServidor
        PortTxt = PuertoDelServidor
    End If
End If

If Winsock1.State <> sckClosed Then
Winsock1.Close
End If

Winsock1.Connect CurServerIp, CurServerPort
'LOCOIP
'127.0.0.1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        prgRun = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Make Server IP and Port box visible
If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    
    'Port
    PortTxt.Visible = True
    'Label4.Visible = True
    
    'Server IP
    PortTxt.Text = "17664"
    IPTxt.Text = "190.0.163.1"
    IPTxt.Visible = True
    'Label5.Visible = True
    
    KeyCode = 0
    Exit Sub
End If

End Sub
Private Sub LoadButtons()
    
    Dim GrhPath As String
    
    GrhPath = DirGraficos
    
    Set cBotonCrearPj = New clsGraphicalButton
    Set cBotonRecuperarPass = New clsGraphicalButton
    Set cBotonManual = New clsGraphicalButton
    Set cBotonReglamento = New clsGraphicalButton
    Set cBotonCodigoFuente = New clsGraphicalButton
    Set cBotonBorrarPj = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    Set cBotonLeerMas = New clsGraphicalButton
    Set cBotonForo = New clsGraphicalButton
    Set cBotonConectarse = New clsGraphicalButton
    Set cBotonTeclas = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton

        
    Call cBotonCrearPj.Initialize(imgCrearPj, GrhPath & "BotonCrearPersonajeConectar.jpg", _
                                    GrhPath & "BotonCrearPersonajeRolloverConectar.jpg", _
                                    GrhPath & "BotonCrearPersonajeClickConectar.jpg", Me)
                                    

    Call cBotonSalir.Initialize(imgSalir, GrhPath & "BotonSalirConnect.jpg", _
                                    GrhPath & "BotonBotonSalirRolloverConnect.jpg", _
                                    GrhPath & "BotonSalirClickConnect.jpg", Me)
                                    
    Call cBotonConectarse.Initialize(imgConectarse, GrhPath & "BotonConectarse.jpg", _
                                    GrhPath & "BotonConectarseRollover.jpg", _
                                    GrhPath & "BotonConectarseClick.jpg", Me)

End Sub

Private Sub imgBorrarPj_Click()

On Error GoTo errH
    Call Shell(App.path & "\RECUPERAR.EXE", vbNormalFocus)

    Exit Sub

errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "Desterium  AO")
End Sub

Private Sub Image1_Click()
Call Audio.PlayWave(SND_CLICK)
MsgBox "Foro en construcción."
End Sub

Private Sub Image2_Click()
Call Audio.PlayWave(SND_CLICK)
ShellExecute Me.hwnd, "open", "http://www.ds-ao.com.ar/", "", "", 1
End Sub

Private Sub Image3_Click()
Call Audio.PlayWave(SND_CLICK)
ShellExecute Me.hwnd, "open", "http://www.ds-ao.com.ar/donaciones", "", "", 1
End Sub

Private Sub Image4_Click()
PanelUser.Show , frmConnect
End Sub

Private Sub imgConectarse_Click()

Call Audio.PlayWave(SND_CLICK)

#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
#Else
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
#End If
    
    'update user info
    UserName = txtNombre.Text
    
    Dim aux As String
    aux = txtPasswd.Text
    
#If SeguridadAlkon Then
    UserPassword = MD5.GetMD5String(aux)
    Call MD5.MD5Reset
#Else
    UserPassword = aux
#End If
    If CheckUserData(False) = True Then
        EstadoLogin = Normal
        
#If UsarWrench = 1 Then
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
#Else
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If

    End If
    
        If checkrecu.value = 1 Then
Dim cantpjs As Byte
    cantpjs = Val(GetVar(App.path & "\Recursos\Datos.DS", "LOG", "CantPjs"))
    WriteVar App.path & "\Recursos\Datos.DS", "PJ" & cantpjs + 1, "Nick", UserName
    WriteVar App.path & "\Recursos\Datos.DS", "PJ" & cantpjs + 1, "Passwd", UserPassword
    WriteVar App.path & "\Recursos\Datos.DS", "LOG", "CantPjs", cantpjs + 1
End If
    
    
End Sub

Private Sub imgCrearPj_Click()
Call Audio.PlayWave(SND_CLICK)
    
    frmConnect.Visible = False
    
    EstadoLogin = E_MODO.Dados
#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
#Else
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If

End Sub

Private Sub imgLeerMas_Click()
    Call ShellExecute(0, "Open", "http://www.ds-ao.com.ar/", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub imgManual_Click()
    Call ShellExecute(0, "Open", "http://www.ds-ao.com.ar/manual", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub imgRecuperar_Click()
On Error GoTo errH

    Call Audio.PlayWave(SND_CLICK)
    Call Shell(App.path & "\RECUPERAR.EXE", vbNormalFocus)
    Exit Sub
errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "Desterium  AO")
End Sub

Private Sub imgReglamento_Click()
    Call ShellExecute(0, "Open", "http://www.ds-ao.com.ar/reglamento", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub imgSalir_Click()
Call Audio.PlayWave(SND_CLICK)
If MsgBox("¿Desea cerrar Desterium  AO?", vbYesNo + vbQuestion, "Desterium  AO") = vbYes Then
  prgRun = False
Else
            Exit Sub
        End If
End Sub

Private Sub imgServArgentina_Click()
    Call Audio.PlayWave(SND_CLICK)
    IPTxt.Text = IPdelServidor
    PortTxt.Text = PuertoDelServidor
End Sub
Private Sub imgVerForo_Click()
    Call ShellExecute(0, "Open", "http://www.ds-ao.com.ar/", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub RECUPERAR_Click()
Call Audio.PlayWave(SND_CLICK)
EstadoLogin = E_MODO.RecuperarPJ
 
#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
#Else
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
 
End Sub
 
Private Sub BORRAR_Click()
Call Audio.PlayWave(SND_CLICK)
EstadoLogin = E_MODO.BorrarPJ
 
#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
#Else
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
 
End Sub

Private Sub txtNombre_Change()
 
Dim F As Long
For F = 1 To MaxRecu
    If UCase$(txtNombre.Text) = UCase$(Recu(F).Nick) Then
        txtPasswd.Text = Recu(F).Password
    End If
Next F
 
End Sub


Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then imgConectarse_Click
End Sub

Private Sub WebAuxiliar_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    
    If InStr(1, URL, "alkon") <> 0 Then
        Call ShellExecute(hwnd, "open", URL, vbNullString, vbNullString, SW_SHOWNORMAL)
        Cancel = True
    End If
    
End Sub

Private Sub webNoticias_NewWindow2(ppDisp As Object, Cancel As Boolean)
    
    WebAuxiliar.RegisterAsBrowser = True
    Set ppDisp = WebAuxiliar.Object
    
End Sub

Private Sub Winsock1_Connect()
OnOff.ForeColor = vbGreen
OnOff.Caption = "Online"
End Sub
 
Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
OnOff.ForeColor = vbRed
OnOff.Caption = "Offline"
End Sub
Private Sub txtPasswd_GotFocus()
          Dim cL As Long
   Dim cantpjs As Byte
   cantpjs = Val(GetVar(App.path & "\Recursos\Datos.DS", "LOG", "CantPjs"))
   For cL = 1 To cantpjs
   If UCase$(GetVar(App.path & "\Recursos\Datos.DS", "PJ" & cL, "NICK")) = UCase$(txtNombre.Text) Then
txtPasswd.Text = GetVar(App.path & "\Recursos\Datos.DS", "PJ" & cL, "Passwd")
End If
Next cL
End Sub
