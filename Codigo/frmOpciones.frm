VERSION 5.00
Begin VB.Form frmOpciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   2350
      MouseIcon       =   "frmOpciones.frx":0152
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   19
      Top             =   3320
      Width           =   335
   End
   Begin VB.PictureBox PictureSanado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   2350
      MouseIcon       =   "frmOpciones.frx":045C
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   7
      Top             =   2880
      Width           =   335
   End
   Begin VB.PictureBox PictureRecuMana 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   840
      MouseIcon       =   "frmOpciones.frx":0766
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   2400
      Width           =   335
   End
   Begin VB.PictureBox PictureVestirse 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   840
      MouseIcon       =   "frmOpciones.frx":0A70
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   5
      Top             =   2880
      Width           =   335
   End
   Begin VB.PictureBox PictureMenosCansado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   2350
      MouseIcon       =   "frmOpciones.frx":0D7A
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   2400
      Width           =   335
   End
   Begin VB.PictureBox PictureNoHayNada 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   2350
      MouseIcon       =   "frmOpciones.frx":1084
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   3
      Top             =   1920
      Width           =   335
   End
   Begin VB.PictureBox PictureOcultarse 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   840
      MouseIcon       =   "frmOpciones.frx":138E
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   2
      Top             =   1920
      Width           =   335
   End
   Begin VB.PictureBox PictureFxs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   2350
      MouseIcon       =   "frmOpciones.frx":1698
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   1200
      Width           =   335
   End
   Begin VB.PictureBox PictureMusica 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   840
      MouseIcon       =   "frmOpciones.frx":19A2
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   1200
      Width           =   335
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00008000&
      Height          =   495
      Left            =   2280
      Top             =   3225
      Width           =   1620
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "Modo ventana"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2715
      TabIndex        =   20
      Top             =   3375
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00008000&
      Height          =   1335
      Left            =   650
      Top             =   1910
      Width           =   3250
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      Height          =   375
      Left            =   650
      Top             =   1185
      Width           =   3250
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Manual de FénixAO"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   720
      MouseIcon       =   "frmOpciones.frx":1CAC
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Has sanado"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2715
      TabIndex        =   17
      Top             =   2940
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "Meditación"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   2460
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "No hay nada aquí"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   2715
      TabIndex        =   15
      Top             =   1980
      Width           =   1140
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Abrigarse"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   2940
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Menos cansado"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2715
      TabIndex        =   13
      Top             =   2470
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Ocultarse"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Mostrar carteles"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Opciones de sonido"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "FXs"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2715
      TabIndex        =   9
      Top             =   1260
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Música"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   1260
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   3840
      MouseIcon       =   "frmOpciones.frx":1FB6
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   615
   End
End
Attribute VB_Name = "frmOpciones"
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
Private Sub Command2_Click()
Me.Visible = False
End Sub

Private Sub Form_Load()


Me.Picture = LoadPicture(DirGraficos & "OpcionesDelJuego.gif")

If Musica = 0 Then
    PictureMusica.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureMusica.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If FX = 0 Then
    PictureFxs.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureFxs.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If NoRes = 1 Then
    Picture1.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    Picture1.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelOcultarse = 1 Then
    PictureOcultarse.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureOcultarse.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelMenosCansado = 1 Then
    PictureMenosCansado.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureMenosCansado.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelVestirse = 1 Then
    PictureVestirse.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureVestirse.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelNoHayNada = 1 Then
    PictureNoHayNada.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureNoHayNada.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelRecuMana = 1 Then
    PictureRecuMana.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureRecuMana.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelSanado = 1 Then
    PictureSanado.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureSanado.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

End Sub
Private Sub Image1_Click()

Me.Visible = False

End Sub
Private Sub Label11_Click()

ShellExecute Me.hwnd, "open", "http://www.fenixao.com.ar/public_html/Html/manual/", "", "", 1

End Sub
Private Sub Picture1_Click()

If NoRes = 0 Then
    NoRes = 1
    Picture1.Picture = LoadPicture(DirGraficos & "tick1.gif")
    Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "ModoVentana", 1)
Else
    NoRes = 0
    Picture1.Picture = LoadPicture(DirGraficos & "tick2.gif")
    Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "ModoVentana", 0)
End If

MsgBox "Este cambio hará efecto recién la próxima vez que ejecutes el juego."

End Sub
Private Sub PictureFxs_Click()

Select Case FX
    Case 0
        FX = 1
        PictureFxs.Picture = LoadPicture(DirGraficos & "tick2.gif")
    Case 1
        FX = 0
        PictureFxs.Picture = LoadPicture(DirGraficos & "tick1.gif")
End Select

End Sub
Private Sub PictureMenosCansado_Click()

If CartelMenosCansado = 0 Then
    CartelMenosCansado = 1
    PictureMenosCansado.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelMenosCansado = 0
    PictureMenosCansado.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "MenosCansado", Str(CartelMenosCansado))

End Sub

Private Sub PictureMusica_Click()

If Not IsPlayingCheck Then
    Musica = 0
    Play_Midi
    PictureMusica.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    Musica = 1
    Stop_Midi
    PictureMusica.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

End Sub

Private Sub PictureNoHayNada_Click()
If CartelNoHayNada = 0 Then
    CartelNoHayNada = 1
    PictureNoHayNada.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelNoHayNada = 0
    PictureNoHayNada.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "NoHayNada", Str(CartelNoHayNada))

End Sub

Private Sub PictureOcultarse_Click()

If CartelOcultarse = 0 Then
    CartelOcultarse = 1
    PictureOcultarse.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelOcultarse = 0
    PictureOcultarse.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Ocultarse", Str(CartelOcultarse))
End Sub

Private Sub PictureRecuMana_Click()
If CartelRecuMana = 0 Then
    CartelRecuMana = 1
    PictureRecuMana.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelRecuMana = 0
    PictureRecuMana.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "RecuMana", Str(CartelRecuMana))

End Sub

Private Sub PictureSanado_Click()
If CartelSanado = 0 Then
    CartelSanado = 1
    PictureSanado.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelSanado = 0
    PictureSanado.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Sanado", Str(CartelSanado))

End Sub

Private Sub PictureVestirse_Click()
If CartelVestirse = 0 Then
    CartelVestirse = 1
    PictureVestirse.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelVestirse = 0
    PictureVestirse.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Vestirse", Str(CartelVestirse))

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving = False And Button = vbLeftButton Then

      DX = X

      dy = Y

      bmoving = True

   End If

   

End Sub

 

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving And ((X <> DX) Or (Y <> dy)) Then

      Move Left + (X - DX), Top + (Y - dy)

   End If

   

End Sub

 

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = vbLeftButton Then

      bmoving = False

   End If

   

End Sub
