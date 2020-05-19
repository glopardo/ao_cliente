VERSION 5.00
Begin VB.Form frmMapa 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   444
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Shape personaje 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000C0&
      Height          =   255
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   4017
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5880
      MouseIcon       =   "frmMapa.frx":0000
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label LabelMapa 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4090
      TabIndex        =   0
      Top             =   560
      Width           =   1335
   End
End
Attribute VB_Name = "frmMapa"
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

Public BotonMapa As Byte
Public MouseX As Long
Public MouseY As Long
Private Sub Form_Click()

If BotonMapa = 2 Then Call TelepPorMapa(MouseX, MouseY)

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

personaje.Left = IzquierdaMapa + (UserPos.X - 50) * 0.18
personaje.Top = TopMapa + (UserPos.Y - 50) * 0.18

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

personaje.Left = IzquierdaMapa + ((UserPos.X - 50) * 0.18)
personaje.Top = TopMapa + ((UserPos.Y - 50) * 0.18)

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

personaje.Left = IzquierdaMapa + (UserPos.X - 50) * 0.18
personaje.Top = TopMapa + (UserPos.Y - 50) * 0.18

End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "MapaJuego.jpg")
End Sub
Private Sub Form_LostFocus()

Me.Visible = False

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BotonMapa = Button

If bmoving = False And Button = vbLeftButton Then
   DX = X
   dy = Y
   bmoving = True
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving And ((X <> DX) Or (Y <> dy)) Then Move Left + (X - DX), Top + (Y - dy)
MouseX = X
MouseY = Y
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub
Private Sub Form_GotFocus()

personaje.Left = IzquierdaMapa + (UserPos.X - 50) * 0.18
personaje.Top = TopMapa + (UserPos.Y - 50) * 0.18

End Sub

Private Sub Image1_Click()
Me.Visible = False
End Sub
