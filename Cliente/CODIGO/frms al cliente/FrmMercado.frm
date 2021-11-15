VERSION 5.00
Begin VB.Form FrmMercado 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   Picture         =   "FrmMercado.frx":0000
   ScaleHeight     =   4515
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image ImgMao 
      Height          =   375
      Left            =   240
      Top             =   720
      Width           =   2535
   End
   Begin VB.Image ImgOfertasRealizadas 
      Height          =   375
      Left            =   240
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Image ImgOfertasRecibidas 
      Height          =   375
      Left            =   360
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Image ImgPublicar 
      Height          =   375
      Left            =   360
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Image ImgQuitar 
      Height          =   375
      Left            =   360
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   360
      Top             =   3600
      Width           =   2415
   End
End
Attribute VB_Name = "FrmMercado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsFormulario As clsFormMovementManager
Public LastPressed As clsGraphicalButton

Public BotonMercado As clsGraphicalButton
Public BotonOfertasHechas As clsGraphicalButton
Public BotonRecibidas As clsGraphicalButton
Public BotonPublicar As clsGraphicalButton
Public BotonQuitar As clsGraphicalButton
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub Form_Load()
    Call LoadButtons
    
        ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image6_Click()
Call Audio.PlayWave(SND_CLICK)
Unload Me
frmMain.SetFocus

End Sub

Private Sub ImgMAO_Click()
    Call Audio.PlayWave(SND_CLICK)
    
    Call WriteRequestMercado
    
End Sub

Private Sub ImgOfertasRealizadas_Click()
    Call Audio.PlayWave(SND_CLICK)
    WriteRequestOfferSentUser
End Sub

Private Sub ImgOfertasRecibidas_Click()
    Call Audio.PlayWave(SND_CLICK)
    WriteRequestOfferUser
    
End Sub

Private Sub ImgPersonajesPublicados_Click()

End Sub

Private Sub ImgPublicar_Click()
    Call Audio.PlayWave(SND_CLICK)
    FrmPublicarMao.Show , frmMain
    Unload Me
End Sub

Private Sub imgQuitar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteQuitarPj
End Sub
Private Sub LoadButtons()
    Dim GrhPath As String
    
    GrhPath = DirGraficos
    
    Set BotonMercado = New clsGraphicalButton
    Set BotonOfertasHechas = New clsGraphicalButton
    Set BotonRecibidas = New clsGraphicalButton
    Set BotonPublicar = New clsGraphicalButton
    Set BotonQuitar = New clsGraphicalButton
    
    
    Set LastPressed = New clsGraphicalButton
    
   ' Call BotonMercado.Initialize(ImgMAO, grhpath & "BotonMAO.jpg", _
                                    grhpath & "BotonMAO1.jpg", _
                                    grhpath & "BotonMAO.jpg", Me)
    
  '  Call BotonOfertasHechas.Initialize(ImgOfertasRealizadas, grhpath & "BotonMisOfertas.jpg", _
                                    grhpath & "BotonMisOfertas1.jpg", _
                                   grhpath & "BotonMisOfertas.jpg", Me)
                                   
   ' Call BotonRecibidas.Initialize(ImgOfertasRecibidas, grhpath & "BotonVerOfertas.jpg", _
                                    grhpath & "BotonVerOfertas1.jpg", _
                                    grhpath & "BotonVerOfertas.jpg", Me)
                                    
    'Call BotonPublicar.Initialize(ImgPublicar, grhpath & "BotonPublicar.jpg", _
                                    grhpath & "BotonPublicar1.jpg", _
                                    grhpath & "BotonPublicar.jpg", Me)
                                    
   ' Call BotonQuitar.Initialize(ImgQuitar, grhpath & "BotonQuitar.jpg", _
                                    grhpath & "BotonQuitar1.jpg", _
                                    grhpath & "BotonQuitar.jpg", Me)
End Sub

