VERSION 5.00
Begin VB.Form frmTransferir 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRepetirNombre 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   350
      Left            =   840
      TabIndex        =   2
      Top             =   1750
      Width           =   1935
   End
   Begin VB.TextBox txtMonto 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   350
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   350
      Left            =   840
      TabIndex        =   0
      Top             =   450
      Width           =   1935
   End
   Begin VB.Image imgCerrar 
      Height          =   255
      Left            =   3240
      Top             =   0
      Width           =   255
   End
   Begin VB.Image imgAceptar 
      Height          =   375
      Left            =   1200
      MousePointer    =   3  'I-Beam
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "frmTransferir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'Me.Picture = LoadPicture(DirGraficos & "Transferir.gif")
End Sub

Private Sub imgAceptar_Click()
    Call SendData("#&" & txtNombre.Text & "|" & txtMonto.Text)
    Unload frmTransferir
End Sub

Private Sub imgCerrar_Click()
    Unload frmTransferir
End Sub

Private Sub txtRepetirNombre_Change()
    cmdAceptar.Enabled = txtNombre.Text = txtRepetirNombre.Text
End Sub
