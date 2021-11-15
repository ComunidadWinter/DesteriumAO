VERSION 5.00
Begin VB.Form FrmCantidad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1500
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   3240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmDrag.frx":0000
   ScaleHeight     =   1500
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
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
      Height          =   375
      Left            =   330
      TabIndex        =   0
      Top             =   478
      Width           =   2625
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   1680
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "FrmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private X As Single
Private Y As Single
Private tX As Byte
Private tY As Byte
Private MouseX As Long
Private MouseY As Long


Private Sub Image1_Click()
 If LenB(FrmCantidad.text1) > 0 Then
        If Not IsNumeric(FrmCantidad.text1) Then Exit Sub  'Should never happen
        If Inventario.sMoveItem Then
        
        If text1 > Inventario.amount(Inventario.SelectedItem) Then
            ShowConsoleMsg "No tienes esa cantidad!", 65, 190, 156, False, False
        Unload Me
        Exit Sub
        End If
        
X = frmMain.MouseX
Y = frmMain.MouseY
  ConvertCPtoTP X, Y, tX, tY
        WriteDragToPos tX, tY, Inventario.SelectedItem, text1
        FrmCantidad.text1.Text = ""
        
        Else
        Call WriteDrop(Inventario.SelectedItem, FrmCantidad.text1.Text)
        FrmCantidad.text1.Text = ""
    End If
    End If
    Unload Me
End Sub

Private Sub Image2_Click()
If Inventario.SelectedItem = 0 Then Exit Sub
    ' If LenB(frmcantidad.Text1) > 0 Then
       ' If Not IsNumeric(frmcantidad.Text1) Then Exit Sub  'Should never happen
        If Inventario.sMoveItem Then 'drag and drop
X = frmMain.MouseX
Y = frmMain.MouseY
  ConvertCPtoTP X, Y, tX, tY
        WriteDragToPos tX, tY, Inventario.SelectedItem, Inventario.amount(Inventario.SelectedItem)
        FrmCantidad.text1.Text = ""
Unload Me
    
Else 'tirar al piso :D
    If Inventario.SelectedItem <> FLAGORO Then
        Call WriteDrop(Inventario.SelectedItem, Inventario.amount(Inventario.SelectedItem))
        Unload Me
    Else
        If UserGLD > 10000 Then
            Call WriteDrop(Inventario.SelectedItem, 10000)
            Unload Me
        Else
            Call WriteDrop(Inventario.SelectedItem, UserGLD)
            Unload Me
        End If
    End If
    End If
    FrmCantidad.text1.Text = ""
    
    Unload Me
End Sub

Private Sub Text1_Change()
On Error GoTo ErrHandler
    If Val(FrmCantidad.text1) < 0 Then
        FrmCantidad.text1 = "1"
    End If
    
        If text1 < 1 Then
    text1 = 1
    End If
    
    If Val(FrmCantidad.text1) > 100000 Then
        FrmCantidad.text1 = "100000"
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    FrmCantidad.text1 = "1"
End Sub


