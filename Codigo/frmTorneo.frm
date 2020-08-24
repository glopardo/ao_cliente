VERSION 5.00
Begin VB.Form frmTorneo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TORNEO - RivendelAO"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   8100
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frm2vs2 
      Caption         =   "2 vs. 2"
      Enabled         =   0   'False
      Height          =   3615
      Left            =   4440
      TabIndex        =   13
      Top             =   240
      Width           =   3495
      Begin VB.ListBox lstParejas 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   3015
      End
      Begin VB.CommandButton cmdEmparejar 
         Caption         =   "Emparejar"
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtPareja 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdElegirGanador 
      Caption         =   "Elegir ganador"
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdAbrirTorneo 
      Caption         =   "Abrir torneo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdComenzarDuelo 
      Caption         =   "Comenzar duelo"
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton cmdSum 
      Caption         =   "Summonear"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   2175
   End
   Begin VB.ListBox lstSiguienteRonda 
      Height          =   1815
      Left            =   2400
      TabIndex        =   6
      Top             =   2880
      Width           =   1815
   End
   Begin VB.ListBox lstDueleando 
      Height          =   450
      Left            =   2400
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtCantParticipantes 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin VB.ListBox lstParticipantes 
      Height          =   2595
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Frame frmTipoTorneo 
      Caption         =   "Tipo de torneo"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton opt2vs2 
         Caption         =   "2 vs. 2"
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton opt1vs1 
         Caption         =   "1 vs. 1"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Label lblCupos 
      Caption         =   "CUPOS LLENOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblParticipanes 
      Caption         =   "Cant. Participantes"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmTorneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbrirTorneo_Click()
    txtCantParticipantes.Enabled = False
End Sub

Private Sub cmdComenzarDuelo_Click()
    Dim jugador1, jugador2 As String
    If lstParticipantes.ListCount > 1 Then
        jugador1 = lstParticipantes.List(0)
        jugador2 = lstParticipantes.List(1)
        
        Call SendData("/TMSG Se enfrentan en esta ronda: " & jugador1 & " Vs. " & jugador2 & ".")
        Call SendData("/TMSG Mucha suerte y que gane el mejor.")
        
        lstParticipantes.RemoveItem (0)
        lstParticipantes.RemoveItem (0)
        
        lstDueleando.AddItem (jugador1)
        lstDueleando.AddItem (jugador2)
        
        'Enviar a los rivales al mapa de duelo
        Call SendData("/TELEP " & jugador1 & " 191 14 41")
        Call SendData("/TELEP " & jugador2 & " 191 33 51")
        Call SendData("INIDT" & jugador1 & "|" & jugador2)
        Call SendData("/CUENTA 5")
        
    End If
End Sub

Private Sub cmdSum_Click()
    Dim i As Integer
    For i = 0 To lstParticipantes.ListCount - 1
        'Enviar a los participantes al mapa de espera
        Call SendData("/TELEP " & lstParticipantes.List(i) & " 200 50 44")
    Next
End Sub

Private Sub opt1vs1_Click()
    opt2vs2.value = Not opt1vs1.value
    frm2vs2.Enabled = Not opt1vs1.value
End Sub

Private Sub opt2vs2_Click()
    opt1vs1.value = Not opt2vs2.value
    frm2vs2.Enabled = opt2vs2.value
End Sub

Private Sub txtCantParticipantes_Change()
    cmdAbrirTorneo.Enabled = txtCantParticipantes.Text <> ""
End Sub
