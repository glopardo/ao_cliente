VERSION 5.00
Begin VB.Form frmMonto 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "x"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtMonto 
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmMonto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Accion As Integer '1=depositar;2=retirar

Private Sub cmdAceptar_Click()
    Select Case Accion
        Case 1
            Call SendData("#Ñ " & txtMonto.Text)
            Unload frmMonto
            Exit Sub
        Case 2
            Call SendData("#0 " & txtMonto.Text)
            Unload frmMonto
            Exit Sub
    End Select
End Sub

Private Sub cmdCancelar_Click()
    Unload frmMonto
End Sub
