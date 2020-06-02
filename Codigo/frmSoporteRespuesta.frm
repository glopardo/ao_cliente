VERSION 5.00
Begin VB.Form frmSoporteRespuesta 
   BorderStyle     =   0  'None
   Caption         =   "Respuesta soporte"
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txtRespuesta 
      Height          =   2895
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmSoporteRespuesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCerrar_Click()
    frmMain.pctEnvelope.Visible = False
    Call SendData("CLSOP")
    Unload frmSoporteRespuesta
End Sub

Private Sub Form_Load()
    txtRespuesta = RespuestaSoporte
End Sub
