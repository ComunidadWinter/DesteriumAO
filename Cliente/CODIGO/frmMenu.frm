VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   Picture         =   "frmMenu.frx":0000
   ScaleHeight     =   4140
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5160
      TabIndex        =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Image ImageQUest 
      Height          =   855
      Left            =   240
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   240
      Top             =   3000
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   240
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Call Audio.PlayWave(SND_CLICK)
Unload Me
FrmMercado.Show vbModeless, frmMain
End Sub

Private Sub Image2_Click()
Call Audio.PlayWave(SND_CLICK)
Unload Me
FrmRanking.Show vbModeless, frmMain
End Sub

Private Sub ImageQUest_Click()
Call Audio.PlayWave(SND_CLICK)
Unload Me
WriteQuestListRequest
End Sub

Private Sub Label1_Click()
Unload Me
End Sub
