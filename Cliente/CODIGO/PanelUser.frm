VERSION 5.00
Begin VB.Form PanelUser 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel de Usario Dolwur"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   2640
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton RECUPERAR 
      Caption         =   "Recuperar Usuario"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton BORRAR 
      Caption         =   "Borrar Usuario"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "PanelUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()
Unload Me

'MsgBox "Atención con esta acción va a eliminar el personaje, no podra volver a usarlo"
End Sub

Private Sub RECUPERAR_Click()
 
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
 Unload Me
End Sub
 
Private Sub BORRAR_Click()
 
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
 Unload Me
End Sub

