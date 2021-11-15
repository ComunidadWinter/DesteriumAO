VERSION 5.00
Begin VB.Form frmPerfil 
   BorderStyle     =   0  'None
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   LinkTopic       =   "Form4"
   Picture         =   "frmPerfil.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picInfo 
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   8580
      Picture         =   "frmPerfil.frx":41306
      ScaleHeight     =   3061.224
      ScaleMode       =   0  'User
      ScaleWidth      =   3000
      TabIndex        =   28
      Top             =   585
      Visible         =   0   'False
      Width           =   3000
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "CLERIGO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   2550
         Left            =   195
         TabIndex        =   29
         Top             =   195
         Width           =   2550
      End
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   3
      Left            =   5900
      Picture         =   "frmPerfil.frx":49527
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   30
      Top             =   2826
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   23
      Left            =   8430
      Picture         =   "frmPerfil.frx":4CAB7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   27
      Top             =   5340
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   22
      Left            =   7605
      Picture         =   "frmPerfil.frx":50047
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   26
      Top             =   5340
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   21
      Left            =   6750
      Picture         =   "frmPerfil.frx":535D7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   25
      Top             =   5340
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   20
      Left            =   5900
      Picture         =   "frmPerfil.frx":56B67
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   24
      Top             =   5340
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   19
      Left            =   5070
      Picture         =   "frmPerfil.frx":5A0F7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   23
      Top             =   5340
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   18
      Left            =   4230
      Picture         =   "frmPerfil.frx":5D687
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   22
      Top             =   5340
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   17
      Left            =   3390
      Picture         =   "frmPerfil.frx":60C17
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   21
      Top             =   5340
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   16
      Left            =   7605
      Picture         =   "frmPerfil.frx":641A7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   20
      Top             =   4485
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   15
      Left            =   6750
      Picture         =   "frmPerfil.frx":67737
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   19
      Top             =   4485
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   14
      Left            =   5900
      Picture         =   "frmPerfil.frx":6ACC7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   18
      Top             =   4485
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   13
      Left            =   5070
      Picture         =   "frmPerfil.frx":6E257
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   17
      Top             =   4485
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   12
      Left            =   4230
      Picture         =   "frmPerfil.frx":717E7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   16
      Top             =   4485
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   11
      Left            =   3390
      Picture         =   "frmPerfil.frx":74D77
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   15
      Top             =   4485
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   10
      Left            =   7590
      Picture         =   "frmPerfil.frx":78307
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   14
      Top             =   3660
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   9
      Left            =   6750
      Picture         =   "frmPerfil.frx":7B897
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   13
      Top             =   3660
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   8
      Left            =   5900
      Picture         =   "frmPerfil.frx":7EE27
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   12
      Top             =   3660
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   7
      Left            =   5070
      Picture         =   "frmPerfil.frx":823B7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   11
      Top             =   3680
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   6
      Left            =   4230
      Picture         =   "frmPerfil.frx":85947
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   10
      Top             =   3660
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   5
      Left            =   3390
      Picture         =   "frmPerfil.frx":88ED7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   9
      Top             =   3660
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   4
      Left            =   6750
      Picture         =   "frmPerfil.frx":8C467
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   8
      Top             =   2826
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   2
      Left            =   5070
      Picture         =   "frmPerfil.frx":8F9F7
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   7
      Top             =   2820
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   745
      Index           =   1
      Left            =   4230
      Picture         =   "frmPerfil.frx":92F87
      ScaleHeight     =   750
      ScaleWidth      =   735
      TabIndex        =   6
      Top             =   2830
      Width           =   740
   End
   Begin VB.PictureBox PicLogro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   0
      Left            =   3390
      Picture         =   "frmPerfil.frx":96517
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   5
      Top             =   2800
      Width           =   735
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   23
      Left            =   8385
      Top             =   5265
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   22
      Left            =   7605
      Top             =   5265
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   21
      Left            =   6630
      Top             =   5265
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   20
      Left            =   5850
      Top             =   5265
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   19
      Left            =   5070
      Top             =   5265
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   18
      Left            =   4290
      Top             =   5265
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   17
      Left            =   3315
      Top             =   5265
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   16
      Left            =   7605
      Top             =   4485
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   15
      Left            =   6825
      Top             =   4485
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   14
      Left            =   5850
      Top             =   4485
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   13
      Left            =   5070
      Top             =   4485
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   12
      Left            =   4290
      Top             =   4485
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   11
      Left            =   3315
      Top             =   4485
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   10
      Left            =   7605
      Top             =   3705
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   9
      Left            =   6825
      Top             =   3705
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   8
      Left            =   5850
      Top             =   3705
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   7
      Left            =   5070
      Top             =   3705
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   6
      Left            =   4290
      Top             =   3705
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   5
      Left            =   3315
      Top             =   3705
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   4
      Left            =   6825
      Top             =   2730
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   3
      Left            =   5850
      Top             =   2730
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   2
      Left            =   5070
      Top             =   2730
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   1
      Left            =   4290
      Top             =   2730
      Width           =   795
   End
   Begin VB.Image ImgInfo 
      Height          =   795
      Index           =   0
      Left            =   3315
      Top             =   2730
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   9165
      Top             =   7800
      Width           =   2160
   End
   Begin VB.Label lblElv 
      BackStyle       =   0  'Transparent
      Caption         =   "47"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   5745
      Width           =   615
   End
   Begin VB.Label lblRaza 
      BackStyle       =   0  'Transparent
      Caption         =   "HUMANO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   5300
      Width           =   1815
   End
   Begin VB.Label lblClase 
      BackStyle       =   0  'Transparent
      Caption         =   "CLERIGO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblClan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<MI GUILD ES CARLO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   3680
      TabIndex        =   1
      Top             =   1890
      Width           =   4455
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MI NOMBRE ES CARLITOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   1200
      Width           =   4455
   End
End
Attribute VB_Name = "frmPerfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picInfo.Visible = False
End Sub
Private Sub ImgInfo_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    picInfo.Visible = True
    picInfo.Move (PicLogro(index).Left + X + 20), (PicLogro(index).Top + Y + 40)
    
    
    InfoRequest index
End Sub
Private Sub PicLogro_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    picInfo.Visible = True
    picInfo.Move (PicLogro(index).Left + X + 20), (PicLogro(index).Top + Y + 40)
    
    
    InfoRequest index
End Sub
Private Sub InfoRequest(ByVal index As Integer)

    Dim tmpStr As String
    
    Select Case index
        Case 0 'Usuario recien registrado
            tmpStr = "REGISTRADO" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado como agradecimiento por jugar Desterium AO" & vbCrLf
            tmpStr = tmpStr & "Recompensa: 10.000 monedas de oro"
            
        Case 1 ' Nivel máximo
            tmpStr = "NIVEL MÁXIMO" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado como recompensa por alcanzar el nivel máximo." & vbCrLf
            tmpStr = tmpStr & "Recompensa: 1.000.000 monedas de oro"
        
        Case 2 ' Fundar clan
            tmpStr = "FUNDADOR" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por haber fundado un clan."
        
        Case 3 ' Oro
            tmpStr = "USUARIO ORO" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado a los usuarios ORO."
            
        Case 4 ' Premium
            tmpStr = "USUARIO PREMIUM" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado a los usuarios PREMIUM."
            
        Case 5 ' Usuarios matados
            tmpStr = "100 FRAGS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 100 Frags" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 250.000 monedas de oro"
        Case 6 ' Usuarios matados
            tmpStr = "200 FRAGS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 200 Frags" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 500.000 monedas de oro"
        Case 7 ' Usuarios matados
            tmpStr = "400 FRAGS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 400 Frags" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 750.000 monedas de oro"
        Case 8 ' Usuarios matados
            tmpStr = "800 FRAGS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 800 Frags" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 1.000.000 monedas de oro"
        Case 9 ' Usuarios matados
            tmpStr = "1600 FRAGS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 1600 Frags" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 2.000.000 monedas de oro"
        Case 10 ' Usuarios matados
            tmpStr = "5000 FRAGS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 5000 Frags" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 4.000.000 monedas de oro y 500 DSP"
        Case 11 ' Retos
            tmpStr = "5 RETOS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 5 Retos GANADOS" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 150.000 monedas de oro "
        Case 12 ' Retos
            tmpStr = "10 RETOS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 10 Retos GANADOS" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 300.000 monedas de oro "
        Case 13 ' Retos
            tmpStr = "50 RETOS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 50 Retos GANADOS" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 500.000 monedas de oro "
        Case 14 ' Retos
            tmpStr = "100 RETOS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 100 Retos GANADOS" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 600.000 monedas de oro "
        Case 15 ' Retos
            tmpStr = "250 RETOS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 250 Retos GANADOS" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 1.000.000 monedas de oro "
        Case 16 ' Retos
            tmpStr = "1000 RETOS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 1000 Retos GANADOS" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 3.500.000 monedas de oro "
        Case 17 ' Eventos
            tmpStr = "5 EVENTOS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 5 Eventos GANADOS" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 300.000 monedas de oro "
        Case 18 ' Eventos
            tmpStr = "10 EVENTOS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 10 Eventos GANADOS" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 500.000 monedas de oro "
        Case 19 ' Eventos
            tmpStr = "20 EVENTOS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 20 Eventos GANADOS" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 600.000 monedas de oro "
        Case 20 ' Eventos
            tmpStr = "30 EVENTOS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 30 Eventos GANADOS" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 700.000 monedas de oro "
        Case 21 ' Eventos
            tmpStr = "40 EVENTOS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 40 Eventos GANADOS" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 800.000 monedas de oro "
        Case 22 ' Eventos
            tmpStr = "50 EVENTOS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 50 Eventos GANADOS" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 900.000 monedas de oro "
        Case 23 ' Eventos
            tmpStr = "100 EVENTOS" & vbCrLf
            tmpStr = tmpStr & "Logro otorgado por alcanzar los 100 Eventos GANADOS" & vbCrLf & vbCrLf
            tmpStr = tmpStr & "Recompensa: 1.000.000 monedas de oro "
    End Select
    
    
    lblInfo.Caption = tmpStr
    
End Sub
'ahora vas a entenderlo dame 2m OKIS
Private Sub picInfo_Click()

End Sub

