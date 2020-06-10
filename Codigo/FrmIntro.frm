VERSION 5.00
Begin VB.Form FrmIntro 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   Icon            =   "FrmIntro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   1680
      MouseIcon       =   "FrmIntro.frx":168B4
      MousePointer    =   99  'Custom
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   1680
      MouseIcon       =   "FrmIntro.frx":16BBE
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Image Image6 
      Height          =   615
      Left            =   2520
      MouseIcon       =   "FrmIntro.frx":16EC8
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   615
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   2520
      MouseIcon       =   "FrmIntro.frx":171D2
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   615
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   2520
      MouseIcon       =   "FrmIntro.frx":174DC
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   615
   End
End
Attribute VB_Name = "FrmIntro"
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

Private Sub Form_Load()

Me.Picture = LoadPicture(App.Path & "\Graficos\MenuRapido.jpg")

End Sub
Private Sub Image2_Click()
    Call Main
End Sub

Private Sub Image3_Click()
    ShellExecute Me.hWnd, "open", App.Path & "/aosetup.exe", "", "", 1
End Sub

Private Sub Image4_Click()
    ShellExecute Me.hWnd, "open", "https://www.instagram.com/", "", "", 1
End Sub

Private Sub Image5_Click()
    ShellExecute Me.hWnd, "open", "http://www.facebook.com", "", "", 1
End Sub

Private Sub Image6_Click()
    ShellExecute Me.hWnd, "open", "https://discord.com/", "", "", 1
End Sub

Private Sub Image7_Click()
    Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If bmoving = False And Button = vbLeftButton Then
      Dx3 = X
      dy = Y
      bmoving = True
   End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If bmoving And ((X <> Dx3) Or (Y <> dy)) Then
      Move left + (X - Dx3), top + (Y - dy)
   End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      bmoving = False
   End If
End Sub


