VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPanelGm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                       Panel GM"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4320
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox R 
      Height          =   270
      Left            =   75
      TabIndex        =   74
      Top             =   4035
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   476
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmPanelGm.frx":0000
   End
   Begin VB.PictureBox SSTab1 
      Height          =   3915
      Left            =   75
      ScaleHeight     =   3855
      ScaleWidth      =   4110
      TabIndex        =   0
      Top             =   45
      Width           =   4170
      Begin VB.CheckBox CheckConsola 
         Caption         =   "No estoy borracho, drogado, etc, y prometo utilizar el rmsg de forma correcta y sin ERRORES DE ORTOGRAFIA!!!!."
         Enabled         =   0   'False
         Height          =   735
         Left            =   -74910
         TabIndex        =   96
         Top             =   1815
         Width           =   3720
      End
      Begin VB.Frame Frame4 
         Caption         =   "Rmsg"
         ForeColor       =   &H00C00000&
         Height          =   1020
         Left            =   -74910
         TabIndex        =   89
         Top             =   2535
         Width           =   3720
         Begin VB.CommandButton Limpiar2 
            Caption         =   "Enviar y Limpiar"
            Enabled         =   0   'False
            Height          =   360
            Left            =   1905
            TabIndex        =   95
            Top             =   570
            Width           =   1725
         End
         Begin VB.CommandButton Enviar2 
            Caption         =   "Enviar"
            Enabled         =   0   'False
            Height          =   360
            Left            =   75
            TabIndex        =   94
            Top             =   570
            Width           =   1830
         End
         Begin VB.TextBox Consola 
            Enabled         =   0   'False
            Height          =   300
            Left            =   60
            TabIndex        =   93
            Top             =   240
            Width           =   3570
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Staff"
         ForeColor       =   &H000000FF&
         Height          =   1005
         Left            =   -74895
         TabIndex        =   88
         Top             =   765
         Width           =   3735
         Begin VB.CommandButton Limpiar 
            Caption         =   "Enviar y Limpiar"
            Enabled         =   0   'False
            Height          =   360
            Left            =   1875
            TabIndex        =   92
            Top             =   555
            Width           =   1755
         End
         Begin VB.CommandButton Enviar 
            Caption         =   "Enviar"
            Enabled         =   0   'False
            Height          =   360
            Left            =   60
            TabIndex        =   91
            Top             =   555
            Width           =   1815
         End
         Begin VB.TextBox Staff 
            Enabled         =   0   'False
            Height          =   300
            Left            =   75
            TabIndex        =   90
            Top             =   225
            Width           =   3570
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Sum / Ira / Telep"
         ForeColor       =   &H00808000&
         Height          =   1425
         Left            =   -74865
         TabIndex        =   76
         Top             =   2310
         Width           =   3720
         Begin VB.CommandButton DONDE 
            Caption         =   "DONDE"
            Height          =   240
            Left            =   30
            TabIndex        =   84
            Top             =   1125
            Width           =   3600
         End
         Begin VB.TextBox Mapa 
            Height          =   240
            Left            =   3030
            TabIndex        =   83
            Top             =   840
            Width           =   570
         End
         Begin VB.TextBox Y 
            Height          =   240
            Left            =   3030
            TabIndex        =   82
            Top             =   555
            Width           =   570
         End
         Begin VB.TextBox X 
            Height          =   240
            Left            =   1920
            TabIndex        =   81
            Top             =   555
            Width           =   630
         End
         Begin VB.CommandButton TELEP 
            Caption         =   "TELEP"
            Height          =   240
            Left            =   60
            TabIndex        =   80
            Top             =   825
            Width           =   2370
         End
         Begin VB.CommandButton IRA 
            Caption         =   "IRA"
            Height          =   255
            Left            =   60
            TabIndex        =   79
            Top             =   540
            Width           =   1560
         End
         Begin VB.TextBox Nick 
            Height          =   240
            Left            =   1680
            TabIndex        =   78
            Text            =   "NICK"
            Top             =   240
            Width           =   1965
         End
         Begin VB.CommandButton SUM 
            Caption         =   "SUM"
            Enabled         =   0   'False
            Height          =   255
            Left            =   75
            TabIndex        =   77
            Top             =   240
            Width           =   1545
         End
         Begin VB.Label Label3 
            Caption         =   "X:"
            Height          =   240
            Left            =   1695
            TabIndex        =   87
            Top             =   570
            Width           =   195
         End
         Begin VB.Label Label2 
            Caption         =   "Y:"
            Height          =   240
            Left            =   2775
            TabIndex        =   86
            Top             =   585
            Width           =   210
         End
         Begin VB.Label Label1 
            Caption         =   "Mapa:"
            Height          =   255
            Left            =   2505
            TabIndex        =   85
            Top             =   855
            Width           =   480
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Left            =   -72780
         TabIndex        =   75
         Top             =   2700
         Width           =   1590
      End
      Begin VB.CommandButton CC 
         Caption         =   "CC"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73425
         TabIndex        =   73
         Top             =   2700
         Width           =   630
      End
      Begin VB.CommandButton Trabajando 
         Caption         =   "Trabajando"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74865
         TabIndex        =   72
         Top             =   2700
         Width           =   1425
      End
      Begin VB.CommandButton Cmap 
         Caption         =   "Cmap"
         Height          =   330
         Left            =   -72330
         TabIndex        =   71
         Top             =   2355
         Width           =   1140
      End
      Begin VB.CommandButton Bloq 
         Caption         =   "Bloquear"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -73350
         TabIndex        =   70
         Top             =   2355
         Width           =   1005
      End
      Begin VB.CommandButton Mata 
         Caption         =   "Mata"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72240
         TabIndex        =   69
         Top             =   2040
         Width           =   1050
      End
      Begin VB.CommandButton Dest 
         Caption         =   "Dest"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74865
         TabIndex        =   68
         Top             =   2040
         Width           =   600
      End
      Begin VB.CommandButton Masskill 
         Caption         =   "Masskill"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73305
         TabIndex        =   67
         Top             =   2040
         Width           =   1035
      End
      Begin VB.CommandButton Massdest 
         Caption         =   "Massdest"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74235
         TabIndex        =   66
         Top             =   2040
         Width           =   915
      End
      Begin VB.CommandButton Omap 
         Caption         =   "Pjs online en mapa"
         Height          =   315
         Left            =   -74865
         TabIndex        =   63
         Top             =   2370
         Width           =   1500
      End
      Begin VB.Frame Frame14 
         Caption         =   "Otros"
         ForeColor       =   &H00C00000&
         Height          =   1590
         Left            =   -74850
         TabIndex        =   53
         Top             =   2175
         Width           =   3690
         Begin VB.Frame Frame15 
            Caption         =   "Broadcast"
            ForeColor       =   &H000000FF&
            Height          =   975
            Left            =   60
            TabIndex        =   56
            Top             =   525
            Width           =   3570
            Begin VB.TextBox Broadcast 
               Enabled         =   0   'False
               Height          =   285
               Left            =   45
               TabIndex        =   58
               Text            =   "Mensaje para el broadcast"
               Top             =   195
               Width           =   3465
            End
            Begin VB.CommandButton ManBroadcast 
               Caption         =   "Mandar broadcast"
               Enabled         =   0   'False
               Height          =   360
               Left            =   75
               TabIndex        =   57
               Top             =   555
               Width           =   3420
            End
         End
         Begin VB.CommandButton ManPausa 
            Caption         =   "Pausar el server"
            Enabled         =   0   'False
            Height          =   285
            Left            =   1890
            TabIndex        =   55
            Top             =   195
            Width           =   1755
         End
         Begin VB.CommandButton ManDats 
            Caption         =   "Reiniciar DATs"
            Enabled         =   0   'False
            Height          =   285
            Left            =   45
            TabIndex        =   54
            Top             =   195
            Width           =   1800
         End
      End
      Begin VB.CommandButton OnlineGM 
         Caption         =   "OnlineGM"
         Height          =   330
         Left            =   -72390
         TabIndex        =   50
         Top             =   1695
         Width           =   1200
      End
      Begin VB.Frame Frame12 
         Caption         =   "Inventarios"
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   -74835
         TabIndex        =   43
         Top             =   1830
         Width           =   3615
         Begin VB.CommandButton VerOro 
            Caption         =   "ORO"
            Enabled         =   0   'False
            Height          =   240
            Left            =   825
            TabIndex        =   49
            Top             =   570
            Width           =   735
         End
         Begin VB.CommandButton InvBoveda 
            Caption         =   "Inventario en boveda"
            Enabled         =   0   'False
            Height          =   285
            Left            =   1695
            TabIndex        =   48
            Top             =   540
            Width           =   1815
         End
         Begin VB.CommandButton VerOroI 
            Caption         =   "OroB"
            Enabled         =   0   'False
            Height          =   255
            Left            =   90
            TabIndex        =   47
            Top             =   555
            Width           =   705
         End
         Begin VB.CommandButton BorrarInvPJ 
            Caption         =   "Borrar inventario del pj"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   90
            TabIndex        =   46
            Top             =   870
            Width           =   3420
         End
         Begin VB.TextBox InvNick 
            Enabled         =   0   'False
            Height          =   255
            Left            =   1680
            TabIndex        =   45
            Text            =   "NICK"
            Top             =   240
            Width           =   1830
         End
         Begin VB.CommandButton VerInventario 
            Caption         =   "Ver inventario"
            Enabled         =   0   'False
            Height          =   270
            Left            =   75
            TabIndex        =   44
            Top             =   240
            Width           =   1500
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Echar"
         ForeColor       =   &H00008000&
         Height          =   645
         Left            =   150
         TabIndex        =   40
         Top             =   3180
         Width           =   3675
         Begin VB.TextBox EcharNick 
            Enabled         =   0   'False
            Height          =   345
            Left            =   1545
            TabIndex        =   42
            Text            =   "NICK"
            Top             =   225
            Width           =   2025
         End
         Begin VB.CommandButton Echar 
            Caption         =   "Echar"
            Enabled         =   0   'False
            Height          =   330
            Left            =   60
            TabIndex        =   41
            Top             =   240
            Width           =   1425
         End
      End
      Begin VB.CommandButton Hora 
         Caption         =   "Hora"
         Height          =   315
         Left            =   -74865
         TabIndex        =   38
         Top             =   1710
         Width           =   630
      End
      Begin VB.CommandButton Lluvia 
         Caption         =   "Parar/Comernzar lluvia"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -74205
         TabIndex        =   37
         Top             =   1695
         Width           =   1785
      End
      Begin VB.Frame Frame10 
         Caption         =   "Carcel"
         ForeColor       =   &H00C000C0&
         Height          =   870
         Left            =   135
         TabIndex        =   32
         Top             =   2295
         Width           =   3720
         Begin VB.TextBox CarcelTiempo 
            Enabled         =   0   'False
            Height          =   240
            Left            =   720
            TabIndex        =   35
            Text            =   "en minutos"
            Top             =   210
            Width           =   1020
         End
         Begin VB.TextBox CarcelNick 
            Enabled         =   0   'False
            Height          =   240
            Left            =   1800
            TabIndex        =   34
            Text            =   "NICK"
            Top             =   210
            Width           =   1785
         End
         Begin VB.CommandButton Encarcelar 
            Caption         =   "Encarcelar"
            Enabled         =   0   'False
            Height          =   240
            Left            =   75
            TabIndex        =   33
            Top             =   525
            Width           =   3525
         End
         Begin VB.Label Label6 
            Caption         =   "Tiempo:"
            Height          =   210
            Left            =   75
            TabIndex        =   36
            Top             =   225
            Width           =   615
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Crear item"
         ForeColor       =   &H00008000&
         Height          =   945
         Left            =   -74865
         TabIndex        =   26
         Top             =   720
         Width           =   3690
         Begin VB.CommandButton Hitem 
            Caption         =   "Crear"
            Enabled         =   0   'False
            Height          =   315
            Left            =   105
            TabIndex        =   29
            Top             =   555
            Width           =   3450
         End
         Begin VB.TextBox Cantidad 
            Enabled         =   0   'False
            Height          =   240
            Left            =   810
            TabIndex        =   28
            Top             =   240
            Width           =   945
         End
         Begin VB.TextBox Objeto 
            Enabled         =   0   'False
            Height          =   240
            Left            =   2400
            TabIndex        =   27
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label4 
            Caption         =   "Cantidad:"
            Height          =   225
            Left            =   105
            TabIndex        =   31
            Top             =   255
            Width           =   720
         End
         Begin VB.Label Label5 
            Caption         =   "Objeto:"
            Height          =   210
            Left            =   1815
            TabIndex        =   30
            Top             =   255
            Width           =   570
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Rastreo"
         ForeColor       =   &H00FF0000&
         Height          =   1470
         Left            =   -74865
         TabIndex        =   21
         Top             =   750
         Width           =   3735
         Begin VB.TextBox TNicksMail 
            Enabled         =   0   'False
            Height          =   240
            Left            =   1770
            TabIndex        =   62
            Text            =   "MAIL"
            Top             =   1125
            Width           =   1905
         End
         Begin VB.CommandButton NicksMail 
            Caption         =   "NICKS"
            Enabled         =   0   'False
            Height          =   240
            Left            =   60
            TabIndex        =   61
            Top             =   1125
            Width           =   1650
         End
         Begin VB.TextBox TNicksIp 
            Enabled         =   0   'False
            Height          =   240
            Left            =   1755
            TabIndex        =   60
            Text            =   "IP"
            Top             =   540
            Width           =   1905
         End
         Begin VB.CommandButton NicksIp 
            Caption         =   "NICKS"
            Enabled         =   0   'False
            Height          =   240
            Left            =   60
            TabIndex        =   59
            Top             =   540
            Width           =   1650
         End
         Begin VB.CommandButton IpNick 
            Caption         =   "IP"
            Enabled         =   0   'False
            Height          =   240
            Left            =   60
            TabIndex        =   25
            Top             =   240
            Width           =   1650
         End
         Begin VB.TextBox TIpNick 
            Enabled         =   0   'False
            Height          =   240
            Left            =   1740
            TabIndex        =   24
            Text            =   "NICK"
            Top             =   240
            Width           =   1905
         End
         Begin VB.CommandButton MailNick 
            Caption         =   "MAIL"
            Enabled         =   0   'False
            Height          =   240
            Left            =   60
            TabIndex        =   23
            Top             =   825
            Width           =   1650
         End
         Begin VB.TextBox TMailNick 
            Enabled         =   0   'False
            Height          =   240
            Left            =   1755
            TabIndex        =   22
            Text            =   "NICK"
            Top             =   840
            Width           =   1905
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Modificar PJ"
         ForeColor       =   &H000080FF&
         Height          =   1050
         Left            =   -74850
         TabIndex        =   16
         Top             =   750
         Width           =   3630
         Begin VB.TextBox ModNick 
            Enabled         =   0   'False
            Height          =   285
            Left            =   45
            TabIndex        =   20
            Text            =   "NICK"
            Top             =   210
            Width           =   1500
         End
         Begin VB.ComboBox Combo3 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1590
            TabIndex        =   19
            Top             =   195
            Width           =   1170
         End
         Begin VB.TextBox ModNum 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2790
            TabIndex        =   18
            Text            =   "NUM"
            Top             =   195
            Width           =   735
         End
         Begin VB.CommandButton ModificarPj 
            Caption         =   "MODIFICAR"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   60
            MaskColor       =   &H000000FF&
            TabIndex        =   17
            Top             =   600
            Width           =   3510
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Mantenimiento"
         ForeColor       =   &H00004080&
         Height          =   1410
         Left            =   -74850
         TabIndex        =   9
         Top             =   735
         Width           =   3705
         Begin VB.CommandButton CambiarWS 
            Caption         =   "Cambiar WS"
            Enabled         =   0   'False
            Height          =   285
            Left            =   30
            TabIndex        =   65
            Top             =   1080
            Width           =   1875
         End
         Begin VB.CommandButton InfoWS 
            Caption         =   "WS echos"
            Enabled         =   0   'False
            Height          =   285
            Left            =   30
            TabIndex        =   64
            Top             =   795
            Width           =   1875
         End
         Begin VB.CommandButton ManBACKUP 
            Caption         =   "Hacer backup completo"
            Enabled         =   0   'False
            Height          =   300
            Left            =   60
            TabIndex        =   15
            Top             =   210
            Width           =   1845
         End
         Begin VB.CommandButton ManGuardar 
            Caption         =   "Guardar pjs"
            Enabled         =   0   'False
            Height          =   285
            Left            =   45
            TabIndex        =   14
            Top             =   510
            Width           =   1860
         End
         Begin VB.CommandButton ManApagar 
            BackColor       =   &H000000FF&
            Caption         =   "Apagar server"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1920
            TabIndex        =   13
            Top             =   1080
            Width           =   1725
         End
         Begin VB.CommandButton ManReiniciar 
            Caption         =   "Reiniciar"
            Enabled         =   0   'False
            Height          =   300
            Left            =   1935
            TabIndex        =   12
            Top             =   210
            Width           =   1710
         End
         Begin VB.CommandButton ManReiniciar1 
            Caption         =   "Reiniciar 1"
            Enabled         =   0   'False
            Height          =   300
            Left            =   1920
            TabIndex        =   11
            Top             =   495
            Width           =   1725
         End
         Begin VB.CommandButton ManReiniciar2 
            Caption         =   "Reiniciar 2"
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            TabIndex        =   10
            Top             =   795
            Width           =   1725
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Bans"
         ForeColor       =   &H000000FF&
         Height          =   1545
         Left            =   135
         TabIndex        =   1
         Top             =   735
         Width           =   3735
         Begin VB.TextBox Causa 
            Enabled         =   0   'False
            Height          =   285
            Left            =   690
            TabIndex        =   51
            Text            =   "CAUSA"
            Top             =   1170
            Width           =   2955
         End
         Begin VB.CommandButton Unban 
            Caption         =   "Unban"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   90
            TabIndex        =   39
            Top             =   675
            Width           =   1590
         End
         Begin VB.CommandButton Ban 
            Caption         =   "Ban"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   8
            Top             =   240
            Width           =   1590
         End
         Begin VB.TextBox BanIP 
            Enabled         =   0   'False
            Height          =   270
            Left            =   1995
            TabIndex        =   7
            Text            =   "IP"
            Top             =   225
            Width           =   1665
         End
         Begin VB.TextBox BanNick 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1995
            TabIndex        =   6
            Text            =   "NICK"
            Top             =   525
            Width           =   1650
         End
         Begin VB.TextBox BanMail 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1995
            TabIndex        =   5
            Text            =   "MAIL"
            Top             =   840
            Width           =   1635
         End
         Begin VB.CheckBox CheckIP 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1755
            TabIndex        =   4
            Top             =   240
            Width           =   195
         End
         Begin VB.CheckBox CheckNick 
            Caption         =   "Check2"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1755
            TabIndex        =   3
            Top             =   540
            Width           =   210
         End
         Begin VB.CheckBox CheckMail 
            Caption         =   "Check3"
            Enabled         =   0   'False
            Height          =   210
            Left            =   1755
            TabIndex        =   2
            Top             =   870
            Width           =   240
         End
         Begin VB.Label Label7 
            Caption         =   "Causa:"
            Height          =   210
            Left            =   135
            TabIndex        =   52
            Top             =   1215
            Width           =   540
         End
      End
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Ban_Click()

If MsgBox("¿Estás seguro?", vbExclamation + vbYesNo, "Panel Gm") = vbYes Then
    If CheckIP.value = vbChecked Then Call SendData("/BANIP " & BanIP.Text)
    If CheckMail.value = vbChecked Then Call SendData("/BANMAIL " & BanMail.Text)
    If CheckNick.value = vbChecked Then Call SendData("/BAN " & Causa.Text & "@" & BanNick.Text)
End If

End Sub

Private Sub Bloq_Click()

Call SendData("/BLOQ")

End Sub

Private Sub BorrarInvPJ_Click()

Call SendData("/RESETINV " & InvNick.Text)

End Sub



Private Sub CC_Click()

Call SendData("/CC")

End Sub

Private Sub CheckConsola_Click()
If CheckConsola.value = vbChecked Then
    Consola.Enabled = True
    Enviar2.Enabled = True
    Limpiar2.Enabled = True
End If

If CheckConsola.value = vbUnchecked Then
    Consola.Enabled = False
    Enviar2.Enabled = False
    Limpiar2.Enabled = False
End If

End Sub

Private Sub Cmap_Click()

Call SendData("/CMAP")

End Sub

Private Sub Dest_Click()

Call SendData("/DEST")

End Sub

Private Sub DONDE_Click()

Call SendData("/DONDE " & Nick.Text)

End Sub

Private Sub Echar_Click()

If MsgBox("¿Estás seguro?", vbExclamation + vbYesNo, "Panel Gm") = vbYes Then
    Call SendData("/ECHAR " & EcharNick.Text)
End If

End Sub

Private Sub Encarcelar_Click()

If MsgBox("¿Estás seguro?", vbExclamation + vbYesNo, "Panel Gm") = vbYes Then
    Call SendData("/CARCEL " & CarcelTiempo.Text & " " & CarcelNick.Text)
End If

End Sub

Private Sub Enviar_Click()

Call SendData("/STAFF " & Staff.Text)

End Sub

Private Sub Enviar2_Click()

Call SendData("/STAFF " & Consola.Text)

End Sub

Private Sub Form_Load()


Select Case VarPrivilegios
    Case Is = 0
        Exit Sub ' Por si las moscas
    Case Is = 2
        VerInventario.Enabled = True
        Trabajando.Enabled = True
        Encarcelar.Enabled = True
        CarcelNick.Enabled = True
        Echar.Enabled = True
        EcharNick.Enabled = True
        Ban.Enabled = True
        Unban.Enabled = True
        BanNick.Enabled = True
        CheckNick.Enabled = True
        SUM.Enabled = True
        CC.Enabled = True
        InvNick.Enabled = True
        VerOroI.Enabled = True
        VerOro.Enabled = True
        InvBoveda.Enabled = True
        BorrarInvPJ.Enabled = True
        Mata.Enabled = True
        OnlineGM.Enabled = True
        IpNick.Enabled = True
        TIpNick.Enabled = True
        NicksIp.Enabled = True
        TNicksIp.Enabled = True
        MailNick.Enabled = True
        TMailNick.Enabled = True
        NicksMail.Enabled = True
        TNicksMail.Enabled = True
        Causa.Enabled = True
        CarcelTiempo.Enabled = True
        CheckConsola.Enabled = True
        Staff.Enabled = True
        Enviar.Enabled = True
        Limpiar.Enabled = True
    Case Is = 3
        VerInventario.Enabled = True
        Trabajando.Enabled = True
        Encarcelar.Enabled = True
        CarcelNick.Enabled = True
        Echar.Enabled = True
        EcharNick.Enabled = True
        Ban.Enabled = True
        Unban.Enabled = True
        BanNick.Enabled = True
        BanIP.Enabled = True
        BanMail.Enabled = True
        CheckNick.Enabled = True
        CheckMail.Enabled = True
        CheckIP.Enabled = True
        SUM.Enabled = True
        CC.Enabled = True
        BorrarInvPJ.Enabled = True
        Mata.Enabled = True
        OnlineGM.Enabled = True
        InfoWS.Enabled = True
        CambiarWS.Enabled = True
        Dest.Enabled = True
        Bloq.Enabled = True
        Masskill.Enabled = True
        Massdest.Enabled = True
        Broadcast.Enabled = True
        ManBroadcast.Enabled = True
        ManApagar.Enabled = True
        ManReiniciar.Enabled = True
        ManReiniciar1.Enabled = True
        ManReiniciar2.Enabled = True
        ManDats.Enabled = True
        ManGuardar.Enabled = True
        ManBACKUP.Enabled = True
        Cantidad.Enabled = True
        Objeto.Enabled = True
        Hitem.Enabled = True
        ModNick.Enabled = True
        ModNum.Enabled = True
        Combo3.Enabled = True
        ModificarPj.Enabled = True
        InvNick.Enabled = True
        VerOroI.Enabled = True
        VerOro.Enabled = True
        InvBoveda.Enabled = True
        BorrarInvPJ.Enabled = True
        Lluvia.Enabled = True
        ManPausa.Enabled = True
        IpNick.Enabled = True
        TIpNick.Enabled = True
        NicksIp.Enabled = True
        TNicksIp.Enabled = True
        MailNick.Enabled = True
        TMailNick.Enabled = True
        NicksMail.Enabled = True
        TNicksMail.Enabled = True
        Causa.Enabled = True
        CarcelTiempo.Enabled = True
        CheckConsola.Enabled = True
        Staff.Enabled = True
        Enviar.Enabled = True
        Limpiar.Enabled = True
End Select


Combo3.AddItem "ORO"
Combo3.AddItem "EXP"
Combo3.AddItem "BODY"
Combo3.AddItem "HEAD"
Combo3.AddItem "CRI"
Combo3.AddItem "CIU"
Combo3.AddItem "HP"
Combo3.AddItem "MAN"
Combo3.AddItem "STA"
Combo3.AddItem "HAM"
Combo3.AddItem "SED"
Combo3.AddItem "ATF"
Combo3.AddItem "ATI"
Combo3.AddItem "ATA"
Combo3.AddItem "ATC"
Combo3.AddItem "ATV"
Combo3.AddItem "LEVEL"

End Sub

Private Sub Hitem_Click()

Call SendData("/HITEM " & Cantidad.Text & " " & Objeto.Text)

End Sub

Private Sub Hora_Click()

Call SendData("/HORA")

End Sub

Private Sub InfoWS_Click()

Call SendData("/INFOWS")

End Sub

Private Sub InvBoveda_Click()

Call SendData("/BOV " & InvNick.Text)

End Sub

Private Sub IpNick_Click()

Call SendData("/IPNICK " & TIpNick.Text)

End Sub

Private Sub IRA_Click()

Call SendData("/IRA " & Nick.Text)

End Sub

Private Sub Limpiar_Click()

Call SendData("/STAFF " & Staff.Text)
Staff.Refresh

End Sub

Private Sub Limpiar2_Click()

Call SendData("/STAFF " & Consola.Text)
Consola.Refresh

End Sub

Private Sub Lluvia_Click()

Call SendData("/LLUVIA")

End Sub

Private Sub MailNick_Click()

Call SendData("/MAILNICK " & TMailNick.Text)

End Sub

Private Sub ManApagar_Click()

If MsgBox("¿ESTÁS SEGURISIMO?", vbExclamation + vbYesNo, "Panel Gm") = vbYes Then
    Call SendData("/APAGAR")
End If

End Sub

Private Sub ManBACKUP_Click()

If MsgBox("¿Estás seguro?", vbExclamation + vbYesNo, "Panel Gm") = vbYes Then
    Call SendData("/DOBACKUP")
End If

End Sub

Private Sub ManBroadcast_Click()

Call SendData("/SMSG " & Broadcast.Text)

End Sub

Private Sub ManDats_Click()

If MsgBox("¿Estás seguro?", vbExclamation + vbYesNo, "Panel Gm") = vbYes Then
    Call SendData("/DATS")
End If

End Sub

Private Sub ManGuardar_Click()

If MsgBox("¿Estás seguro?", vbExclamation + vbYesNo, "Panel Gm") = vbYes Then
    Call SendData("/GRABAR")
End If

End Sub

Private Sub ManPausa_Click()

If MsgBox("¿Estás seguro?", vbExclamation + vbYesNo, "Panel Gm") = vbYes Then
    Call SendData("/PAUSA")
End If

End Sub

Private Sub ManReiniciar_Click()

If MsgBox("¿Estás seguro?", vbExclamation + vbYesNo, "Panel Gm") = vbYes Then
    Call SendData("/REINICIAR")
End If

End Sub

Private Sub ManReiniciar1_Click()

If MsgBox("¿Estás seguro?", vbExclamation + vbYesNo, "Panel Gm") = vbYes Then
    Call SendData("/REINICIAR1")
End If

End Sub

Private Sub ManReiniciar2_Click()

If MsgBox("¿Estás seguro?", vbExclamation + vbYesNo, "Panel Gm") = vbYes Then
    Call SendData("/REINICIAR2")
End If

End Sub

Private Sub Masdest_Click()

Call SendData("/MASSDEST")

End Sub

Private Sub Masskill_Click()

Call SendData("/MASSKILL")

End Sub

Private Sub Mata_Click()

Call SendData("/MATA")

End Sub

Private Sub ModificarPj_Click()

Dim Seleccionado As String
Seleccionado = Combo3.Text

Call SendData("/MOD " & ModNick.Text & " " & Seleccionado & " " & ModNum.Text)

End Sub

Private Sub NicksIp_Click()

Call SendData("/NICKSIP " & TNicksIp.Text)

End Sub

Private Sub NicksMail_Click()

Call SendData("/NICKSMAIL " & TNicksMail.Text)

End Sub

Private Sub Omap_Click()

Call SendData("/OMAP")

End Sub

Private Sub OnlineGM_Click()

If VarPrivilegios = 3 Then
    Call SendData("/SONLINE")
Else
    Call SendData("/ONLINEGM")
End If

End Sub

Private Sub r_Change()

R.SelStart = Len(R.Text)

End Sub

Private Sub SUM_Click()

Call SendData("/SUM " & Nick.Text)

End Sub

Private Sub TELEP_Click()

Call SendData("/TELEP " & Nick.Text & " " & Mapa.Text & " " & X.Text & " " & Y.Text)

End Sub

Private Sub Trabajando_Click()

Call SendData("/TRABAJANDO")

End Sub

Private Sub Unban_Click()

If MsgBox("¿Estás seguro?", vbExclamation + vbYesNo, "Panel Gm") = vbYes Then
    If CheckIP.value = vbChecked Then Call SendData("/UNBANIP " & BanIP.Text)
    If CheckMail.value = vbChecked Then Call SendData("/UNBANMAIL " & BanMail.Text)
    If CheckNick.value = vbChecked Then Call SendData("/UNBAN " & BanNick.Text)
End If

End Sub

Private Sub VerInventario_Click()

Call SendData("/INV " & InvNick.Text)

End Sub
