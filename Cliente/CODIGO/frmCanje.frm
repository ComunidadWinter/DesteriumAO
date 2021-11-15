VERSION 5.00
Begin VB.Form frmCanje 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Canje"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3200
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "canjear"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmCanje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
WriteCanjear List1.ListIndex + 1
End Sub

Private Sub Form_Load()
Dim i As Long
For i = 1 To NumCanjes
With tCanje(i)
If .ObjName <> vbNullString Then
 List1.AddItem .ObjName
End If
End With
Next i
End Sub

Private Sub List1_Click()
Label1.Caption = tCanje(List1.ListIndex + 1).PointsR
Dim tGrh As Integer
tGrh = tCanje(List1.ListIndex).GrhIndex
Dim src As RECT, destr As RECT
    With src
        .Left = GrhData(tGrh).sX
        .Top = GrhData(tGrh).sY
        .Right = .Left + GrhData(tGrh).pixelWidth
        .Bottom = .Top + GrhData(tGrh).pixelHeight
    End With
   
    With destr
        .Left = 0
        .Top = 0
        .Right = 32
        .Bottom = 32
    End With
DrawGrhtoHdc Picture1.hdc, tGrh, src, destr
End Sub
