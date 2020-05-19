VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   0  'User
   ScaleWidth      =   12075.47
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCorreo2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      Left            =   345
      TabIndex        =   31
      Top             =   3240
      Width           =   4080
   End
   Begin VB.TextBox txtPasswdCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   345
      PasswordChar    =   "*"
      TabIndex        =   33
      Top             =   4470
      Width           =   4080
   End
   Begin VB.TextBox txtPasswd 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   345
      PasswordChar    =   "*"
      TabIndex        =   32
      Top             =   3840
      Width           =   4080
   End
   Begin VB.TextBox txtCorreo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      Left            =   345
      TabIndex        =   30
      Top             =   2625
      Width           =   4080
   End
   Begin VB.ComboBox lstGenero 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonajedados.frx":0000
      Left            =   5160
      List            =   "frmCrearPersonajedados.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   1950
      Width           =   2753
   End
   Begin VB.ComboBox lstRaza 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonajedados.frx":001D
      Left            =   5160
      List            =   "frmCrearPersonajedados.frx":0030
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   1350
      Width           =   2753
   End
   Begin VB.ComboBox lstHogar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonajedados.frx":005D
      Left            =   5160
      List            =   "frmCrearPersonajedados.frx":0070
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   2580
      Width           =   2753
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   300
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   0
      Top             =   780
      Width           =   3615
   End
   Begin VB.Label modCarisma 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6795
      TabIndex        =   47
      Top             =   7800
      Width           =   690
   End
   Begin VB.Label modInteligencia 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6720
      TabIndex        =   46
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label modConstitucion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6795
      TabIndex        =   45
      Top             =   5850
      Width           =   690
   End
   Begin VB.Label modAgilidad 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6795
      TabIndex        =   44
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label modfuerza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6797
      TabIndex        =   43
      Top             =   3900
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   960
      MouseIcon       =   "frmCrearPersonajedados.frx":00A1
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label lblPass2OK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   4560
      TabIndex        =   42
      Top             =   3750
      Width           =   345
   End
   Begin VB.Label lbSabiduria 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "+3"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   180
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblMailOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   4560
      TabIndex        =   38
      Top             =   2520
      Width           =   345
   End
   Begin VB.Label lblMail2OK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   4560
      TabIndex        =   36
      Top             =   3150
      Width           =   345
   End
   Begin VB.Label lblPassOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   4560
      TabIndex        =   34
      Top             =   4350
      Width           =   345
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   21
      Left            =   11031
      TabIndex        =   29
      Top             =   7200
      Width           =   405
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   42
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":03AB
      MousePointer    =   99  'Custom
      Top             =   7290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   43
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":04FD
      MousePointer    =   99  'Custom
      Top             =   7305
      Width           =   195
   End
   Begin VB.Label puntosquedan 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "32"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6300
      TabIndex        =   28
      Top             =   2955
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   8865
      TabIndex        =   27
      Top             =   525
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   3
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":064F
      MousePointer    =   99  'Custom
      Top             =   1185
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   5
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":07A1
      MousePointer    =   99  'Custom
      Top             =   1440
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   7
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":08F3
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   9
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":0A45
      MousePointer    =   99  'Custom
      Top             =   2070
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   11
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":0B97
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   13
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":0CE9
      MousePointer    =   99  'Custom
      Top             =   2700
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   15
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":0E3B
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   17
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":0F8D
      MousePointer    =   99  'Custom
      Top             =   3270
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   19
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":10DF
      MousePointer    =   99  'Custom
      Top             =   3615
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   21
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":1231
      MousePointer    =   99  'Custom
      Top             =   3945
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   23
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":1383
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   25
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":14D5
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   27
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":1627
      MousePointer    =   99  'Custom
      Top             =   4815
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   1
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":1779
      MousePointer    =   99  'Custom
      Top             =   840
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":18CB
      MousePointer    =   99  'Custom
      Top             =   870
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   2
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":1A1D
      MousePointer    =   99  'Custom
      Top             =   1200
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":1B6F
      MousePointer    =   99  'Custom
      Top             =   1500
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   6
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":1CC1
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   8
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":1E13
      MousePointer    =   99  'Custom
      Top             =   2085
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":1F65
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":20B7
      MousePointer    =   99  'Custom
      Top             =   2730
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   14
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":2209
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   16
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":235B
      MousePointer    =   99  'Custom
      Top             =   3360
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   18
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":24AD
      MousePointer    =   99  'Custom
      Top             =   3630
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":25FF
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   22
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":2751
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":28A3
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   26
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":29F5
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   28
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":2B47
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   29
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":2C99
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":2DEB
      MousePointer    =   99  'Custom
      Top             =   5490
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   31
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":2F3D
      MousePointer    =   99  'Custom
      Top             =   5430
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":308F
      MousePointer    =   99  'Custom
      Top             =   5760
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":31E1
      MousePointer    =   99  'Custom
      Top             =   5760
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":3333
      MousePointer    =   99  'Custom
      Top             =   6105
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   35
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":3485
      MousePointer    =   99  'Custom
      Top             =   6090
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   36
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":35D7
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   37
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":3729
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   38
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":387B
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   39
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":39CD
      MousePointer    =   99  'Custom
      Top             =   6705
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":3B1F
      MousePointer    =   99  'Custom
      Top             =   6990
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":3C71
      MousePointer    =   99  'Custom
      Top             =   6990
      Width           =   135
   End
   Begin VB.Image boton 
      Height          =   255
      Index           =   1
      Left            =   120
      MouseIcon       =   "frmCrearPersonajedados.frx":3DC3
      MousePointer    =   99  'Custom
      Top             =   8640
      Width           =   1125
   End
   Begin VB.Image boton 
      Appearance      =   0  'Flat
      Height          =   450
      Index           =   0
      Left            =   360
      MouseIcon       =   "frmCrearPersonajedados.frx":3F15
      MousePointer    =   99  'Custom
      Top             =   7800
      Width           =   4560
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   20
      Left            =   11031
      TabIndex        =   26
      Top             =   6900
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   19
      Left            =   11031
      TabIndex        =   25
      Top             =   6600
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   18
      Left            =   11031
      TabIndex        =   24
      Top             =   6285
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   17
      Left            =   11031
      TabIndex        =   23
      Top             =   5970
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   16
      Left            =   11031
      TabIndex        =   22
      Top             =   5685
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   15
      Left            =   11031
      TabIndex        =   21
      Top             =   5385
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   14
      Left            =   11031
      TabIndex        =   20
      Top             =   5070
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   13
      Left            =   11031
      TabIndex        =   19
      Top             =   4770
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   12
      Left            =   11031
      TabIndex        =   18
      Top             =   4470
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   11
      Left            =   11031
      TabIndex        =   17
      Top             =   4155
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   10
      Left            =   11031
      TabIndex        =   16
      Top             =   3840
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   9
      Left            =   11031
      TabIndex        =   15
      Top             =   3540
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   8
      Left            =   11031
      TabIndex        =   14
      Top             =   3225
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   7
      Left            =   11031
      TabIndex        =   13
      Top             =   2925
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   6
      Left            =   11031
      TabIndex        =   12
      Top             =   2610
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   5
      Left            =   11031
      TabIndex        =   11
      Top             =   2310
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   4
      Left            =   11031
      TabIndex        =   10
      Top             =   2010
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   3
      Left            =   11031
      TabIndex        =   9
      Top             =   1710
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   2
      Left            =   11031
      TabIndex        =   8
      Top             =   1395
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   0
      Left            =   11031
      TabIndex        =   7
      Top             =   780
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   1
      Left            =   11031
      TabIndex        =   6
      Top             =   1080
      Width           =   398
   End
   Begin VB.Label lbCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   6315
      TabIndex        =   5
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   6315
      TabIndex        =   4
      Top             =   6750
      Width           =   495
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   6315
      TabIndex        =   3
      Top             =   5730
      Width           =   495
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   6315
      TabIndex        =   2
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label lbFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   6315
      TabIndex        =   1
      Top             =   3780
      Width           =   495
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FénixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar
Option Explicit

Public SkillPoints As Byte
Function CheckData() As Boolean

If UserRaza = 0 Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserHogar = 0 Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

If UserSexo = -1 Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If SkillPoints > 0 Then
    MsgBox "Asigne los skillpoints del personaje."
    Exit Function
End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    If UserAtributos(i) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next i

CheckData = True

End Function
Private Sub boton_Click(Index As Integer)
Dim i As Integer
Dim k As Object
        
Call PlayWaveDS(SND_CLICK)

Select Case Index
    Case 0
        LlegoConfirmacion = False
        Confirmacion = 0

        i = 1
        
        For Each k In Skill
            UserSkills(i) = k.Caption
            i = i + 1
        Next
        
        UserName = txtNombre.Text
        
        If Right$(UserName, 1) = " " Then
            UserName = Trim(UserName)
            MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
        
        UserRaza = lstRaza.ListIndex + 1
        UserSexo = lstGenero.ListIndex
        UserHogar = lstHogar.ListIndex + 1
        
        UserAtributos(1) = 1
        UserAtributos(2) = 1
        UserAtributos(3) = 1
        UserAtributos(4) = 1
        UserAtributos(5) = 1
        
        If CheckData() Then
            UserPassword = MD5String(txtPasswd.Text)
            UserEmail = txtCorreo.Text
            
            If Not CheckMailString(UserEmail) Then
                MsgBox "Direccion de mail inválida.", vbExclamation, "Fenix AO"
                txtCorreo.SetFocus
                Exit Sub
            End If
    
            If UserEmail <> txtCorreo2.Text Then
                MsgBox "Las direcciones de mail no coinciden.", vbExclamation, "Fenix AO"
                txtCorreo2.Text = ""
                txtCorreo2.SetFocus
                Exit Sub
            End If
            
            If Len(Trim(txtPasswd)) = 0 Then
                MsgBox "Tenés que ingresar una contraseña.", vbExclamation, "Fenix AO"
                txtPasswd.SetFocus
                Exit Sub
            End If
            
            If Len(Trim(txtPasswd)) < 6 Then
                MsgBox "El password debe tener al menos 6 caracteres.", vbExclamation, "Fenix AO"
                txtPasswd = ""
                txtPasswdCheck = ""
                txtPasswd.SetFocus
                Exit Sub
            End If
            
            If Trim(txtPasswd) <> Trim(txtPasswdCheck) Then
                MsgBox "Las contraseñas no coinciden.", vbInformation, "Fenix AO"
                txtPasswd = ""
                txtPasswdCheck = ""
                txtPasswd.SetFocus
                Exit Sub
            End If
    
            frmMain.Socket1.HostName = IPdelServidor
            frmMain.Socket1.RemotePort = PuertoDelServidor
    
            Me.MousePointer = 11
            EstadoLogin = CrearNuevoPj
    
            If Not frmMain.Socket1.Connected Then
                Call MsgBox("Error: Se ha perdido la conexion con el server.")
                Unload Me
            Else
                Call Login(ValidarLoginMSG(CInt(bRK)))
            End If
            
            If Musica = 0 Then
                CurMidi = DirMidi & "2.mid"
                LoopMidi = 1
                Call CargarMIDI(CurMidi)
                Call Play_Midi
            End If
        
            frmConnect.Picture = LoadPicture(App.Path & "\Graficos\conectar.jpg")
        End If

    Case 1
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
End Select

End Sub
Private Sub command1_Click(Index As Integer)
Call PlayWaveDS(SND_CLICK)

Dim indice
If Index Mod 2 = 0 Then
    If SkillPoints > 0 Then
        indice = Index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = Index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

puntos.Caption = SkillPoints
End Sub
Private Sub Form_Load()

SkillPoints = 10
puntos.Caption = SkillPoints
Me.Picture = LoadPicture(App.Path & "\graficos\CrearPersonajeConDados.gif")
Me.MousePointer = vbDefault

Select Case (lstRaza.List(lstRaza.ListIndex))
    Case Is = "Humano"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = "+ 2"
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = ""
        modCarisma.Caption = ""
    Case Is = "Elfo"
        modfuerza.Caption = ""
        modConstitucion.Caption = "+ 1"
        modAgilidad.Caption = "+ 3"
        modInteligencia.Caption = "+ 1"
        modCarisma.Caption = "+ 2"
    Case Is = "Elfo Oscuro"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = "+ 2"
        modCarisma.Caption = "- 3"
    Case Is = "Enano"
        modfuerza.Caption = "+ 3"
        modConstitucion.Caption = "+ 3"
        modAgilidad.Caption = "- 1"
        modInteligencia.Caption = "- 6"
        modCarisma.Caption = "- 3"
    Case Is = "Gnomo"
        modfuerza.Caption = "- 5"
        modAgilidad.Caption = "+ 4"
        modInteligencia.Caption = "+ 3"
        modCarisma.Caption = "+ 1"
End Select

End Sub

Private Sub Pîcture4_Click()

End Sub

Private Sub Image1_Click()
PlayWaveDS (SND_CLICK)
Call SendData("TIRDAD")
End Sub

Private Sub lstRaza_click()

Select Case (lstRaza.List(lstRaza.ListIndex))
    Case Is = "Humano"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = "+ 2"
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = ""
        modCarisma.Caption = ""
    Case Is = "Elfo"
        modfuerza.Caption = ""
        modConstitucion.Caption = "+ 1"
        modAgilidad.Caption = "+ 3"
        modInteligencia.Caption = "+ 1"
        modCarisma.Caption = "+ 2"
    Case Is = "Elfo Oscuro"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = "+ 2"
        modCarisma.Caption = "- 3"
    Case Is = "Enano"
        modfuerza.Caption = "+ 3"
        modConstitucion.Caption = "+ 3"
        modAgilidad.Caption = "- 1"
        modInteligencia.Caption = "- 6"
        modCarisma.Caption = "- 3"
    Case Is = "Gnomo"
        modfuerza.Caption = "- 5"
        modAgilidad.Caption = "+ 4"
        modInteligencia.Caption = "+ 3"
        modCarisma.Caption = "+ 1"
End Select

End Sub
Private Sub txtCorreo_Change()

If Not CheckMailString(txtCorreo) Then
    lblMailOK = "O"
    lblMailOK.ForeColor = &HC0&
    lblMail2OK = "O"
    lblMail2OK.ForeColor = &HC0&
    Exit Sub
End If

lblMailOK = "P"
lblMailOK.ForeColor = &H80FF&

If (UCase$(txtCorreo.Text) = UCase$(txtCorreo2.Text)) Then
    lblMail2OK = "P"
    lblMail2OK.ForeColor = &H80FF&
Else
    lblMail2OK = "O"
    lblMail2OK.ForeColor = &HC0&
End If

End Sub
Private Sub txtCorreo_GotFocus()

MsgBox "La dirección de correo electrónico DEBE SER real."

End Sub
Private Sub txtCorreo2_Change()

If Not CheckMailString(txtCorreo) Then
    lblMailOK = "O"
    lblMailOK.ForeColor = &HC0&
    lblMail2OK = "O"
    lblMail2OK.ForeColor = &HC0&
    Exit Sub
End If

lblMailOK = "P"
lblMailOK.ForeColor = &H80FF&

If (UCase$(txtCorreo.Text) = UCase$(txtCorreo2.Text)) Then
    lblMail2OK = "P"
    lblMail2OK.ForeColor = &H80FF&
Else
    lblMail2OK = "O"
    lblMail2OK.ForeColor = &HC0&
End If

End Sub
Private Sub txtPasswd_Change()

If Len(Trim(txtPasswd)) < 6 Then
    lblPass2OK = "O"
    lblPass2OK.ForeColor = &HC0&
    lblPassOK = "O"
    lblPassOK.ForeColor = &HC0&
    Exit Sub
End If

lblPass2OK = "P"
lblPass2OK.ForeColor = &H80FF&

If (txtPasswdCheck = txtPasswd) Then
    lblPassOK = "P"
    lblPassOK.ForeColor = &H80FF&
Else
    lblPassOK = "O"
    lblPassOK.ForeColor = &HC0&
End If

End Sub
Private Sub txtPasswdCheck_Change()

If Len(Trim(txtPasswd)) < 6 Then
    lblPass2OK = "O"
    lblPass2OK.ForeColor = &HC0&
    lblPassOK = "O"
    lblPassOK.ForeColor = &HC0&
    Exit Sub
End If

lblPass2OK = "P"
lblPass2OK.ForeColor = &H80FF&

If (txtPasswdCheck = txtPasswd) Then
    lblPassOK = "P"
    lblPassOK.ForeColor = &H80FF&
Else
    lblPassOK = "O"
    lblPassOK.ForeColor = &HC0&
End If

End Sub
Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
MsgBox "Sea cuidadoso al seleccionar el nombre de su personaje, Argentum es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotación politica los administradores borrarán su personaje y no habrá ninguna posibilidad de recuperarlo."
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub
