VERSION 5.00
Begin VB.Form frmRetos 
   BorderStyle     =   0  'None
   Caption         =   "Retos"
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRechazar 
      Caption         =   "Rechazar"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblDesc 
      Caption         =   "quiere retarte a un duelo"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblContrincante 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "frmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAceptar_Click()
    Call SendData("ACPRE")
    Unload frmRetos
End Sub

Private Sub btnRechazar_Click()
    Call SendData("RECRE")
    Unload frmRetos
End Sub
