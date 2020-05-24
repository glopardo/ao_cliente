VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{B370EF78-425C-11D1-9A28-004033CA9316}#2.0#0"; "Captura.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   11580
      Top             =   0
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.PictureBox Minimap 
      AutoRedraw      =   -1  'True
      Height          =   1500
      Left            =   6960
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   38
      Top             =   720
      Width           =   1500
   End
   Begin VB.PictureBox Inventario 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   9000
      ScaleHeight     =   176
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   37
      Top             =   2520
      Width           =   2400
   End
   Begin RichTextLib.RichTextBox rectxt 
      Height          =   1065
      Left            =   240
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   840
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1879
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frInvent 
      BorderStyle     =   0  'None
      Height          =   4358
      Left            =   8520
      TabIndex        =   18
      Top             =   1680
      Width           =   3216
      Begin VB.Image Image5 
         Height          =   195
         Index           =   3
         Left            =   1515
         MouseIcon       =   "frmMain.frx":0949
         MousePointer    =   99  'Custom
         Top             =   3920
         Width           =   255
      End
      Begin VB.Image Image5 
         Height          =   195
         Index           =   2
         Left            =   1515
         MouseIcon       =   "frmMain.frx":0C53
         MousePointer    =   99  'Custom
         Top             =   3600
         Width           =   255
      End
      Begin VB.Image Image5 
         Height          =   255
         Index           =   1
         Left            =   1730
         MouseIcon       =   "frmMain.frx":0F5D
         MousePointer    =   99  'Custom
         Top             =   3700
         Width           =   200
      End
      Begin VB.Image Image5 
         Height          =   255
         Index           =   0
         Left            =   1380
         MouseIcon       =   "frmMain.frx":1267
         MousePointer    =   99  'Custom
         Top             =   3705
         Width           =   195
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   480
         Left            =   3240
         Top             =   3480
         Width           =   480
      End
      Begin VB.Label lblHechizos 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1800
         MouseIcon       =   "frmMain.frx":1571
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   0
         Width           =   1080
      End
      Begin VB.Image imgFondoInvent 
         Height          =   4395
         Left            =   0
         Top             =   0
         Width           =   3240
      End
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1230
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1980
      Visible         =   0   'False
      Width           =   7027
   End
   Begin VB.Frame frHechizos 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   4395
      Left            =   8520
      TabIndex        =   20
      Top             =   1680
      Width           =   3240
      Begin VB.ListBox lstHechizos 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2565
         Left            =   420
         TabIndex        =   21
         Top             =   1095
         Width           =   2595
      End
      Begin VB.Label lblInvent 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   405
         MouseIcon       =   "frmMain.frx":187B
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   15
         Width           =   1350
      End
      Begin VB.Label lblLanzar 
         BackStyle       =   0  'Transparent
         Height          =   360
         Left            =   390
         MouseIcon       =   "frmMain.frx":1B85
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   3840
         Width           =   1305
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Height          =   360
         Left            =   1965
         MouseIcon       =   "frmMain.frx":1E8F
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   3840
         Width           =   1050
      End
      Begin VB.Label lblAbajo 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
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
         Left            =   2760
         MouseIcon       =   "frmMain.frx":2199
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   780
         Width           =   300
      End
      Begin VB.Label lblArriba 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
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
         Left            =   2415
         MouseIcon       =   "frmMain.frx":24A3
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   780
         Width           =   300
      End
      Begin VB.Image imgFondoHechizos 
         Height          =   4395
         Left            =   0
         Top             =   0
         Width           =   3240
      End
   End
   Begin VB.PictureBox Renderer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6240
      Left            =   180
      ScaleHeight     =   416
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   541
      TabIndex        =   36
      Top             =   2295
      Width           =   8115
      Begin Captura.wndCaptura Captura1 
         Left            =   0
         Top             =   0
         _ExtentX        =   688
         _ExtentY        =   688
      End
   End
   Begin VB.Image Image1 
      Height          =   420
      Index           =   3
      Left            =   8640
      MouseIcon       =   "frmMain.frx":27AD
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   3285
   End
   Begin VB.Image Party 
      Height          =   300
      Left            =   10485
      MouseIcon       =   "frmMain.frx":2AB7
      MousePointer    =   99  'Custom
      Top             =   7860
      Width           =   1170
   End
   Begin VB.Label NumOnline 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   210
      Left            =   5909
      TabIndex        =   35
      Top             =   8625
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10920
      TabIndex        =   34
      Top             =   900
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   11160
      TabIndex        =   33
      Top             =   900
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6720
      TabIndex        =   32
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10680
      TabIndex        =   31
      Top             =   900
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label modo 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "1 Normal"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   30
      Top             =   1980
      Width           =   750
   End
   Begin VB.Image Image9 
      Height          =   300
      Left            =   7680
      Top             =   8589
      Width           =   180
   End
   Begin VB.Image Image8 
      Height          =   285
      Left            =   7080
      Top             =   8605
      Width           =   225
   End
   Begin VB.Label Agilidad 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   7920
      TabIndex        =   29
      Top             =   8625
      Width           =   225
   End
   Begin VB.Label Fuerza 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   7320
      TabIndex        =   28
      Top             =   8625
      Width           =   300
   End
   Begin VB.Label casco 
      BackColor       =   &H00000000&
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   3245
      TabIndex        =   1
      Top             =   8625
      Width           =   540
   End
   Begin VB.Label armadura 
      BackColor       =   &H00000000&
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   521
      TabIndex        =   17
      Top             =   8625
      Width           =   540
   End
   Begin VB.Label escudo 
      BackColor       =   &H00000000&
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   2338
      TabIndex        =   16
      Top             =   8640
      Width           =   540
   End
   Begin VB.Label arma 
      BackColor       =   &H00000000&
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1444
      TabIndex        =   15
      Top             =   8625
      Width           =   540
   End
   Begin VB.Label mapa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ullathorpe"
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
      Height          =   375
      Left            =   8640
      TabIndex        =   14
      Top             =   8640
      Width           =   3015
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   10440
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label cantidadhp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   9000
      TabIndex        =   12
      Top             =   7253
      Width           =   1095
   End
   Begin VB.Label cantidadagua 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   9000
      TabIndex        =   11
      Top             =   7954
      Width           =   1095
   End
   Begin VB.Label cantidadsta 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   9000
      TabIndex        =   13
      Top             =   6615
      Width           =   1095
   End
   Begin VB.Label cantidadhambre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   9000
      TabIndex        =   10
      Top             =   7620
      Width           =   1095
   End
   Begin VB.Label cantidadmana 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   9000
      TabIndex        =   9
      Top             =   6963
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   8760
      MouseIcon       =   "frmMain.frx":2DC1
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00003E25&
      X1              =   16
      X2              =   551.467
      Y1              =   126.333
      Y2              =   126.333
   End
   Begin VB.Image Image3 
      Height          =   405
      Left            =   10680
      MouseIcon       =   "frmMain.frx":30CB
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label fpstext 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   720
      TabIndex        =   8
      Top             =   90
      Width           =   45
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Thusing"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8880
      TabIndex        =   7
      Top             =   675
      Width           =   2625
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H00008080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   75
      Left            =   8836
      Top             =   6675
      Width           =   1290
   End
   Begin VB.Shape MANShp 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   75
      Left            =   8835
      Top             =   7020
      Width           =   1290
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   10800
      TabIndex        =   6
      Top             =   6615
      Width           =   90
   End
   Begin VB.Shape Hpshp 
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   8836
      Top             =   7320
      Width           =   1290
   End
   Begin VB.Shape COMIDAsp 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   75
      Left            =   8836
      Top             =   7680
      Width           =   1290
   End
   Begin VB.Shape AGUAsp 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   75
      Left            =   8836
      Top             =   7995
      Width           =   1290
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   0
      Left            =   10485
      MouseIcon       =   "frmMain.frx":33D5
      MousePointer    =   99  'Custom
      Top             =   6840
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   1
      Left            =   10485
      MouseIcon       =   "frmMain.frx":36DF
      MousePointer    =   99  'Custom
      Top             =   7200
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   2
      Left            =   10485
      MouseIcon       =   "frmMain.frx":39E9
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   1245
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
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
      Height          =   225
      Left            =   8880
      TabIndex        =   5
      Top             =   916
      Width           =   465
   End
   Begin VB.Label LvlLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 (52,32%)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   225
      Left            =   9360
      TabIndex        =   4
      Top             =   916
      Width           =   1050
   End
   Begin VB.Label exp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp:"
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
      Height          =   225
      Left            =   8880
      TabIndex        =   3
      Top             =   1105
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10440
      TabIndex        =   2
      Top             =   916
      Visible         =   0   'False
      Width           =   105
   End
End
Attribute VB_Name = "frmMain"
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

Public tX As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long
Public MouseIX As Integer
Public MouseIY As Integer
Public BotonIClick As Integer


Public IsPlaying As Byte
Public boton As Integer

Private Sub Form_Activate()

If frmParty.Visible Then frmParty.SetFocus
If frmParty2.Visible Then frmParty2.SetFocus

End Sub

Private Sub Renderer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

boton = Button

End Sub

Private Sub Image5_Click(Index As Integer)

If (ItemElegido <= 0 Or ItemElegido > MAX_INVENTORY_SLOTS) Then Exit Sub
If ItemElegido = 1 And Index = 0 Then Exit Sub
If ItemElegido = MAX_INVENTORY_SLOTS And Index = 1 Then Exit Sub
If ItemElegido < 6 And Index = 2 Then Exit Sub
If ItemElegido > MAX_INVENTORY_SLOTS - 5 And Index = 3 Then Exit Sub

Call SendData("ZI" & ItemElegido & "," & Index)

Select Case Index
    Case 0
        ItemElegido = ItemElegido - 1
    Case 1
        ItemElegido = ItemElegido + 1
    Case 2
        ItemElegido = ItemElegido - 5
    Case 3
        ItemElegido = ItemElegido + 5
End Select

End Sub

Private Sub Image7_Click()

End Sub

Private Sub Label3_Click()

Call SendData("#N")

End Sub

Private Sub Label5_Click()

Call SendData("#!")

End Sub

Private Sub Label7_Click()

Call SendData("#O")

End Sub

Private Sub lblarriba_Click()

If lstHechizos.ListIndex < 1 Then Exit Sub

If lstHechizos.ListIndex >= 1 Then Call SendData("DESPHE" & 1 & "," & lstHechizos.ListIndex + 1)
lstHechizos.ListIndex = lstHechizos.ListIndex - 1

End Sub
Private Sub lblabajo_Click()

If lstHechizos.ListIndex > 33 Then Exit Sub

If lstHechizos.ListIndex <= 33 Then Call SendData("DESPHE" & 2 & "," & lstHechizos.ListIndex + 1)
lstHechizos.ListIndex = lstHechizos.ListIndex + 1

End Sub
Private Sub Renderer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

MouseX = X
MouseY = Y

End Sub
Private Sub FX_Timer()
Dim n As Byte

If FX = 0 And RandomNumber(1, 150) < 12 Then
    n = RandomNumber(1, 45)
    Select Case n
        Case Is <= 15
            Call Audio.PlayWave("22.wav")
        Case Is <= 30
            Call Audio.PlayWave("21.wav")
        Case Is <= 35
            Call Audio.PlayWave("28.wav")
        Case Is <= 40
            Call Audio.PlayWave("29.wav")
        Case Is <= 45
            Call Audio.PlayWave("34.wav")
    End Select
End If

End Sub

Private Sub imgObjeto_DblClick(Index As Integer)

If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub

If ItemElegido = Index Then Call SendData("USE" & ItemElegido)
    
End Sub
Private Sub lblHechizos_Click()

Call Audio.PlayWave(SND_CLICK)
frHechizos.Visible = True
frInvent.Visible = False
Inventario.Visible = False

End Sub
Private Sub lblInvent_Click()

Call Audio.PlayWave(SND_CLICK)
frInvent.Visible = True
frHechizos.Visible = False
Inventario.Visible = True
ActualizarInv = True
End Sub

Private Sub lblObjCant_DblClick(Index As Integer)

If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub

If ItemElegido = Index Then Call SendData("USE" & ItemElegido)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If prgRun Then
    prgRun = False
    Cancel = 1
End If

End Sub
Private Sub Image2_Click()

Me.WindowState = vbMinimized

End Sub
Private Sub Image4_Click()

ItemElegido = FLAGORO
If UserGLD > 0 Then frmCantidad.Show

End Sub
Private Sub Party_Click()

frmParty.ListaIntegrantes.Clear
LlegoParty = False
Call SendData("PARINF")
Do While Not LlegoParty
    DoEvents
Loop
frmParty.Visible = True
frmParty.SetFocus
LlegoParty = False
            
End Sub
Private Sub RecTxt_GotFocus()

SendTxt.Visible = False
frmMain.SetFocus

End Sub
Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Call ProcesaEntradaCmd(stxtbuffer)
    stxtbuffer = ""
    frmMain.SendTxt.Text = ""
    frmMain.SendTxt.Visible = False
    KeyCode = 0
End If

End Sub
Private Sub TirarItem()
    If (ItemElegido > 0 And ItemElegido < MAX_INVENTORY_SLOTS + 1) Or (ItemElegido = FLAGORO) Then
        If UserInventory(ItemElegido).Amount = 1 Then
            SendData "TI" & ItemElegido & "," & 1
        Else
           If UserInventory(ItemElegido).Amount > 1 Then
            frmCantidad.Show
           End If
        End If
    End If

 
End Sub

Private Sub AgarrarItem()
    SendData "AG"
 
End Sub

Private Sub UsarItem()
    If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then
    SendData "USA" & ItemElegido
    End If
   
End Sub
Public Sub EquiparItem()

If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then _
        SendData "EQUI" & ItemElegido
        
End Sub
Private Sub lblLanzar_Click()

If lstHechizos.List(lstHechizos.ListIndex) <> "Nada" And TiempoTranscurrido(LastHechizo) >= IntervaloSpell And TiempoTranscurrido(Hechi) >= IntervaloSpell / 4 Then
    Call SendData("LH" & lstHechizos.ListIndex + 1)
    Call SendData("UK" & Magia)
End If

End Sub
Private Sub lblInfo_Click()
    Call SendData("INFS" & lstHechizos.ListIndex + 1)
End Sub
Private Sub Renderer_Click()

If Cartel Then Cartel = False

If Comerciando = 0 Then
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    If Abs(UserPos.Y - tY) > 6 Then Exit Sub
    If Abs(UserPos.X - tX) > 8 Then Exit Sub
    If EligiendoWhispereo Then
        Call SendData("WH" & tX & "," & tY)
        EligiendoWhispereo = False
        Exit Sub
    End If
    
    If UsingSkill = 0 Then
        SendData "LC" & tX & "," & tY
    Else
        frmMain.MousePointer = vbDefault
        If UsingSkill = Magia Then
            If (TiempoTranscurrido(LastHechizo) < IntervaloSpell Or TiempoTranscurrido(Hechi) < IntervaloSpell / 4) Then
                Exit Sub
            Else: Hechi = Timer
            End If
        ElseIf UsingSkill = Proyectiles Then
            If (TiempoTranscurrido(LastFlecha) < IntervaloFlecha Or TiempoTranscurrido(Flecho) < IntervaloFlecha / 4) Then
                Exit Sub
            Else: Flecho = Timer
            End If
        End If
        Call SendData("WLC" & tX & "," & tY & "," & UsingSkill)
        UsingSkill = 0
    End If
End If

If boton = vbRightButton Then Call SendData("/TELEPLOC")
boton = 0

End Sub

Private Sub Renderer_DblClick()
    If Not frmForo.Visible Then
        SendData "RC" & tX & "," & tY
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If Not SendTxt.Visible Then

    Select Case KeyCode

            
        Case vbKeyM:
            Select Case Musica
                Case 0
                    Musica = 1
                    Audio.StopMidi
                    frmOpciones.PictureMusica.Picture = LoadPicture(DirGraficos & "tick2.gif")
                Case 1
                    Musica = 0
                    Audio.PlayMIDI CurMidi
                    frmOpciones.PictureMusica.Picture = LoadPicture(DirGraficos & "tick1.gif")
            End Select
            Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "Musica", Trim(Str(Musica)))
            
        Case vbKeyA:
            Call AgarrarItem
        Case vbKeyC:
        
        
        Case vbKeyE:
            Call EquiparItem
        Case vbKeyF12:
            Nombres = Not Nombres
        Case vbKeyZ:
            Call SendData("(A")
            
        Case vbKeyD
            Call SendData("UK" & Domar)
            
        Case vbKeyR:
            Call SendData("UK" & Robar)
    
        Case vbKeyO:
            Call SendData("UK" & Ocultarse)
            
        Case vbKeyT:
            Call TirarItem
            
        Case vbKey1:
            frmMain.modo = "1 Normal"
            If EligiendoWhispereo Then
                EligiendoWhispereo = False
                MousePointer = 1
            End If
            
        Case vbKey2:
            Call AddtoRichTextBox(frmMain.rectxt, "Has click sobre el usuario al que quieres susurrar.", 255, 255, 255, 1, 0)
            frmMain.modo = "2 Susurrar"
            MousePointer = 2
            EligiendoWhispereo = True
            
        Case vbKey3:
            frmMain.modo = "3 Clan"
            If EligiendoWhispereo Then
                EligiendoWhispereo = False
                MousePointer = 1
            End If

        Case vbKey4:
            frmMain.modo = "4 Grito"
            If EligiendoWhispereo Then
                EligiendoWhispereo = False
                MousePointer = 1
            End If
            
        Case vbKey5:
            frmMain.modo = "5 Rol"
            If EligiendoWhispereo Then
                EligiendoWhispereo = False
                MousePointer = 1
            End If
        
        Case vbKey6:
            frmMain.modo = "6 Party"
            If EligiendoWhispereo Then
                EligiendoWhispereo = False
                MousePointer = 1
            End If
             
        Case vbKeyU:
            If Not NoPuedeUsar Then
                NoPuedeUsar = True
                Call UsarItem
            End If
            
        Case vbKeyL:
            Call SendData("RPU")
            Beep
        
        Case Else
            If vigilar And KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And KeyCode <> vbKeyLeft And KeyCode <> vbKeyRight And KeyCode <> vbKeyReturn And KeyCode <> vbKeyF7 And KeyCode <> vbKeyF1 And KeyCode <> vbKeyF5 And KeyCode <> vbKeyControl Then SendData "FRF" & Chr(KeyCode)
    
    End Select
End If

Select Case KeyCode
    Case vbKeyReturn:
        If Not frmCantidad.Visible Then
            SendTxt.Visible = True
            SendTxt.SetFocus
        End If
        
    Case vbKeyF3:
        Call SendData("/INVISIBLE")
        
    Case vbKeyF5:
        Dim i As Integer
        Captura1.Area = Ventana
        Captura1.Captura
        For i = 1 To 1000
            If Not FileExist(App.Path & "\screenshots\Imagen" & i & ".bmp", vbNormal) Then Exit For
        Next
        Call SavePicture(Captura1.Imagen, App.Path & "/screenshots/Imagen" & i & ".bmp")
        Call AddtoRichTextBox(frmMain.rectxt, "Una imagen fue guardada en la carpeta de screenshots bajo el nombre de Imagen" & i & ".bmp", 255, 150, 50, False, False, False)
    
    Case vbKeyF7:
        Call SendData("/MEDITAR")
    
    Case vbKeyF9:
        frmParty.ListaIntegrantes.Clear
        LlegoParty = False
        Call SendData("PARINF")
        Do While Not LlegoParty
            DoEvents
        Loop
        frmParty.Visible = True
        frmParty.SetFocus
        LlegoParty = False
    
    Case vbKeyControl:
        If (TiempoTranscurrido(LastGolpe) >= IntervaloGolpe) And (TiempoTranscurrido(Golpeo) >= IntervaloGolpe / 4) And (Not UserDescansar) And _
           (Not UserMeditar) Then
            Call SendData("AT")
            Golpeo = Timer
        End If
              
End Select

End Sub
Sub Form_Load()
'BETA
IPdelServidor = "186.19.73.65"
PuertoDelServidor = 10200

FPSFLAG = True

Me.Picture = LoadPicture(DirGraficos & "Principal.gif")
Image8.Picture = LoadPicture(DirGraficos & "Verde.gif")
Image9.Picture = LoadPicture(DirGraficos & "Amarilla.gif")

frmMain.imgFondoInvent.Picture = LoadPicture(DirGraficos & "Centronuevoinventario.gif")
frmMain.imgFondoHechizos.Picture = LoadPicture(DirGraficos & "Centronuevohechizos.gif")

End Sub
Private Sub lstHechizos_KeyDown(KeyCode As Integer, Shift As Integer)

KeyCode = 0

End Sub
Private Sub lstHechizos_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub
Private Sub lstHechizos_KeyUp(KeyCode As Integer, Shift As Integer)

KeyCode = 0

End Sub
Private Sub Image1_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0
        Call frmOpciones.Show(vbModeless, frmMain)
    Case 1
        LlegaronAtrib = False
        LlegaronSkills = False
        LlegoFama = False
        LlegoMinist = False
        SendData "ATRI"
        SendData "ESKI"
        SendData "FAMA"
        Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama Or Not LlegoMinist
            DoEvents
        Loop
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Show
        LlegaronAtrib = False
        LlegaronSkills = False
        LlegoFama = False
        LlegoMinist = False
    Case 2
        If frmGuildLeader.Visible Then frmGuildLeader.Visible = False
        If frmGuildsNuevo.Visible Then frmGuildsNuevo.Visible = False
        If frmGuildAdm.Visible Then frmGuildAdm.Visible = False
        Call SendData("GLINFO")
    Case 3
       frmMapa.Visible = True
End Select

End Sub

Private Sub Image3_Click()
frmSalir.Show


End Sub

Private Sub Label1_Click()
LlegaronSkills = False
SendData "ESKI"

Do While Not LlegaronSkills
    DoEvents
Loop

Dim i As Integer
For i = 1 To NUMSKILLS
    frmSkills3.Text1(i).Caption = UserSkills(i)
Next i
Alocados = SkillPoints
frmSkills3.Puntos.Caption = SkillPoints
frmSkills3.Show
End Sub
Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mx As Integer
Dim my As Integer
Dim aux As Integer
mx = X \ 32 + 1
my = Y \ 32 + 1
aux = (mx + (my - 1) * 5) + OffsetDelInv

End Sub

Private Sub Inventario_DblClick()
 
Dim X As Integer
Dim Y As Integer
 
X = (MouseIX + 16) / 32
Y = (MouseIY + 16) / 32
 
ItemElegido = (Y - 1) * 5 + X
 
If ItemElegido < 1 Then ItemElegido = 1
If ItemElegido > 25 Then ItemElegido = 25
 
If BotonIClick = 2 Then
Call SendData("EQUI" & ItemElegido)
Else
Call SendData("USE" & ItemElegido) ': pocionesCount = pocionesCount + 1
End If
 
End Sub
 
Private Sub Inventario_Click()

ActualizarInv = True

Dim X As Integer
Dim Y As Integer
 
X = (MouseIX + 16) / 32
Y = (MouseIY + 16) / 32
 
ItemElegido = (Y - 1) * 5 + X
 
If BotonIClick = 2 Then Call SendData("EQUI" & ItemElegido)
 
End Sub
 
Private Sub Inventario_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
MouseIX = X
MouseIY = Y
 
End Sub
 
Private Sub Inventario_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseIX = X
MouseIY = Y
BotonIClick = Button


 
End Sub

Private Sub RecTxt_Change()
On Error Resume Next

If SendTxt.Visible Then
    SendTxt.SetFocus
ElseIf (Not frmComerciar.Visible) And _
    (Not frmSkills3.Visible) And _
    (Not frmMSG.Visible) And _
    (Not frmForo.Visible) And _
    (Not frmEstadisticas.Visible) And _
    (Not frmCantidad.Visible) Then
      ' Picture1.SetFocus
End If

End Sub
Private Sub SendTxt_Change()

stxtbuffer = SendTxt.Text
    
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0
          
End Sub
Private Sub Socket1_Connect()
    
    If EstadoLogin = CrearNuevoPj Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = Normal Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = dados Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = Activar Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = RecuperarPAss Then
            Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = BorrarPJ Then
            Call SendData("gIvEmEvAlcOde")
    End If
End Sub


Private Sub Socket1_Disconnect()
    logged = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    
    frmMain.Visible = False

    Pausa = False
    UserMeditar = False

    UserSexo = 0
    UserRaza = 0
    UserEmail = ""
    bO = 100
    
    Dim i As Integer
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

End Sub
Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)

Select Case ErrorCode
    Case 24036
        Call MsgBox("Por favor espere, intentando completar conexión.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub

    Case 24038, 24061
        Call MsgBox("No se puede establecer la conexión con el servidor.", vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")

    Case 24053
        Call MsgBox("Conexión perdida.", vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")
        
    Case 24060
        Call MsgBox("Tiempo de espera agotado.", vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")
    
    Case Else
        Call MsgBox(ErrorString, vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")
     
End Select

frmConnect.MousePointer = 1
Response = 0

frmMain.Socket1.Disconnect

If Not frmCrearPersonaje.Visible Then
    frmConnect.Show
Else
    frmCrearPersonaje.MousePointer = 0
End If

End Sub
Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
Dim loopc As Integer

Dim RD As String
Dim rBuffer(1 To 500) As String

Static TempString As String

Dim CR As Integer
Dim tChar As String
Dim sChar As Integer

Call Socket1.Read(RD, DataLength)

If TempString <> "" Then
    RD = TempString & RD
    TempString = ""
End If

sChar = 1

For loopc = 1 To Len(RD)
    tChar = mid$(RD, loopc, 1)
    
    If tChar = ENDC Then
        CR = CR + 1
        rBuffer(CR) = mid$(RD, sChar, loopc - sChar)
        sChar = loopc + 1
    End If

Next loopc

If Len(RD) - (sChar - 1) <> 0 Then TempString = mid$(RD, sChar, Len(RD))

For loopc = 1 To CR
    Call HandleData(rBuffer(loopc))
Next loopc

End Sub
