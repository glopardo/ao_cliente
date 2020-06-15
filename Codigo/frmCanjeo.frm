VERSION 5.00
Begin VB.Form frmCanjeo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPuntosDisponibles 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtPrecio 
      BackColor       =   &H80000008&
      ForeColor       =   &H0000FF00&
      Height          =   405
      Left            =   4080
      TabIndex        =   6
      Top             =   5880
      Width           =   855
   End
   Begin VB.ListBox lstDesc 
      Height          =   1815
      Left            =   3120
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   550
      Left            =   5280
      ScaleHeight     =   811.364
      ScaleMode       =   0  'User
      ScaleWidth      =   525
      TabIndex        =   4
      Top             =   1080
      Width           =   550
   End
   Begin VB.CommandButton cmdCanjear 
      Caption         =   "Canjear"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   6720
      Width           =   3375
   End
   Begin VB.TextBox txtItemDesc 
      Height          =   3735
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   2895
   End
   Begin VB.ListBox lstItems 
      Columns         =   1
      Height          =   4935
      ItemData        =   "frmCanjeo.frx":0000
      Left            =   240
      List            =   "frmCanjeo.frx":0002
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label lblPrecio 
      Caption         =   "puntos de canje"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   6000
      Width           =   1335
   End
End
Attribute VB_Name = "frmCanjeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PuntosCanje As Long
Dim ArrayItemObjIndex(0 To 99) As Integer
Dim ArrayItemGrhIndex(0 To 99) As Integer
Dim ArrayItemDescripcion(0 To 99) As String
Dim ArrayItemPuntosCanje(0 To 99) As Integer

Private Sub cmdCanjear_Click()
    If PuntosCanje >= ArrayItemPuntosCanje(lstItems.ListIndex) Then
        Call SendData("FINCANJE" & ArrayItemObjIndex(lstItems.ListIndex) & "|" & ArrayItemPuntosCanje(lstItems.ListIndex))
        Unload frmCanjeo
        Exit Sub
    Else
        AddtoRichTextBox frmMain.rectxt, "No tenés suficientes puntos de canje.", 2, 51, 223, 1, 1
    End If
End Sub

Private Sub cmdCerrar_Click()
    Unload frmCanjeo
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    txtPuntosDisponibles.Text = PuntosCanje
    'lstDesc: desc|index|grh
    
    For i = 0 To lstItems.ListCount - 1
        lstDesc.ListIndex = i
        ArrayItemObjIndex(i) = ReadField(1, lstDesc.Text, 124)
        ArrayItemDescripcion(i) = ReadField(2, lstDesc.Text, 124)
        ArrayItemPuntosCanje(i) = CInt(ReadField(3, lstDesc.Text, 124))
        ArrayItemGrhIndex(i) = ReadField(4, lstDesc.Text, 124)
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase ArrayItemGrhIndex
    Erase ArrayItemDescripcion
End Sub

Private Sub lstItems_Click()
    txtItemDesc.Text = ArrayItemDescripcion(lstItems.ListIndex)
    Call DrawGrhtoHdc(Picture1.hDC, ArrayItemGrhIndex(lstItems.ListIndex))
    txtPrecio.Text = ArrayItemPuntosCanje(lstItems.ListIndex)
    cmdCanjear.Enabled = PuntosCanje >= ArrayItemPuntosCanje(lstItems.ListIndex)
End Sub
