VERSION 5.00
Begin VB.Form frmactivar 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox codigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1190
      TabIndex        =   2
      Top             =   2870
      Width           =   3015
   End
   Begin VB.TextBox pass 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1180
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1907
      Width           =   3015
   End
   Begin VB.TextBox nombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   950
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   1080
      MouseIcon       =   "activar.frx":0000
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Image volver 
      Height          =   375
      Left            =   0
      MouseIcon       =   "activar.frx":030A
      MousePointer    =   99  'Custom
      Top             =   4320
      Width           =   1095
   End
End
Attribute VB_Name = "frmactivar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "ActivarPersonaje.gif")

End Sub
Private Sub Image1_Click()

EstadoLogin = Activar
frmMain.Socket1.HostName = IPdelServidor
frmMain.Socket1.RemotePort = PuertoDelServidor
Me.MousePointer = 11
frmMain.Socket1.Connect

End Sub
Private Sub Image2_Click()

Call SendData("DESACT" & nombre.Text & "," & pass.Text & "," & codigo.Text)

End Sub
Private Sub volver_Click()

If Musica = 0 Then
    CurMidi = DirMidi & "2.mid"
    LoopMidi = 1
    Call CargarMIDI(CurMidi)
    Call Play_Midi
End If

frmConnect.Picture = LoadPicture(App.Path & "\Graficos\conectar.jpg")

frmMain.Socket1.Disconnect
frmConnect.MousePointer = 1
Unload Me
      
End Sub
