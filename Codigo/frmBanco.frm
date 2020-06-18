VERSION 5.00
Begin VB.Form frmBanco 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRetirar 
      Caption         =   "Retirar"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdDepositar 
      Caption         =   "Depositar"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdTransferir 
      Caption         =   "Transferir oro"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdBoveda 
      Caption         =   "Abrir bóveda"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "X"
      Height          =   255
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBoveda_Click()
    Call SendData("#%")
End Sub

Private Sub cmdCerrar_Click()
    Unload frmBanco
End Sub

Private Sub cmdDepositar_Click()
    frmMonto.Accion = 1
    frmMonto.Show
End Sub

Private Sub cmdRetirar_Click()
    frmMonto.Accion = 2
    frmMonto.Show
End Sub

Private Sub cmdTransferir_Click()
    frmTransferir.Show
End Sub
