VERSION 5.00
Begin VB.Form frmBancoObj 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   485
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox cantidad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004DC488&
      Height          =   285
      Left            =   3720
      TabIndex        =   7
      Text            =   "1"
      Top             =   6645
      Width           =   600
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   720
      ScaleHeight     =   570
      ScaleWidth      =   525
      TabIndex        =   2
      Top             =   720
      Width           =   555
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004DC488&
      Height          =   3930
      Index           =   1
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Width           =   2490
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004DC488&
      Height          =   3930
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   2040
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   2
      Left            =   4080
      TabIndex        =   9
      Top             =   1050
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   8
      Top             =   1335
      Width           =   45
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   6120
      MouseIcon       =   "frmBancoObj.frx":0000
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   1
      Left            =   3840
      MouseIcon       =   "frmBancoObj.frx":030A
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6165
      Width           =   2460
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   0
      Left            =   615
      MouseIcon       =   "frmBancoObj.frx":0614
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6150
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   3
      Left            =   4080
      TabIndex        =   6
      Top             =   1500
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   4
      Left            =   4080
      TabIndex        =   5
      Top             =   1275
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   5
      Left            =   1320
      TabIndex        =   4
      Top             =   1650
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   1042
      Width           =   45
   End
End
Attribute VB_Name = "frmBancoObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez



Option Explicit

'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^ y
'   le puse borde a la ventana.
'
'[END]'

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
Public LastIndex1 As Integer
Public LastIndex2 As Integer




Private Sub cantidad_Change()
If Val(cantidad.Text) < 0 Then
    cantidad.Text = 1
End If

If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
    cantidad.Text = 1
End If

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Command2_Click()

End Sub



Private Sub Form_Deactivate()

Me.SetFocus

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving = False And Button = vbLeftButton Then
   DX = X
   dy = Y
   bmoving = True
End If

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub

Private Sub Form_Load()
'Cargamos la interfase
frmBancoObj.Picture = LoadPicture(App.Path & "\Graficos\Boveda.jpg")

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Image1(0).Tag = 0 Then
    Image1(0).Tag = 1
End If

If Image1(1).Tag = 0 Then
    Image1(1).Tag = 1
End If

If bmoving And ((X <> DX) Or (Y <> dy)) Then Move Left + (X - DX), Top + (Y - dy)

End Sub

Private Sub Image1_Click(Index As Integer)

Call PlayWaveDS(SND_CLICK)

If List1(Index).List(List1(Index).ListIndex) = "Nada" Or _
   List1(Index).ListIndex < 0 Then Exit Sub

Select Case Index
    Case 0
        frmBancoObj.List1(0).SetFocus
        LastIndex1 = List1(0).ListIndex
        Lista = 0
        SendData ("RETI" & "," & List1(0).ListIndex + 1 & "," & cantidad.Text)
                
   Case 1
        LastIndex2 = List1(1).ListIndex
        If UserInventory(List1(1).ListIndex + 1).Equipped = 0 Then
            Lista = 1
            SendData ("DEPO" & "," & List1(1).ListIndex + 1 & "," & cantidad.Text)
        Else
            AddtoRichTextBox frmMain.rectxt, "No podes depositar el item porque lo estás usando.", 2, 51, 223, 1, 1
            Exit Sub
        End If
                
End Select

NPCInvDim = 0
End Sub
Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
    Case 0
        If Image1(0).Tag = 1 Then
            Image1(0).Tag = 0
            Image1(1).Tag = 1
        End If
        
    Case 1
        If Image1(1).Tag = 1 Then
            Image1(1).Tag = 0
            Image1(0).Tag = 1
        End If
        
End Select

End Sub
Private Sub Image2_Click()
SendData ("FINBAN")
End Sub
Private Sub List1_Click(Index As Integer)

Lista = Index
Call ActualizarInformacionBoveda(Index)

End Sub
Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeyE:
        If List1(1).ListIndex > -1 And List1(1).ListIndex < MAX_INVENTORY_SLOTS - 1 Then
            Call SendData("EQUI" & List1(1).ListIndex + 1)
        End If
End Select

End Sub

Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Image1(0).Tag = 0 Then Image1(0).Tag = 1
If Image1(1).Tag = 0 Then Image1(1).Tag = 1

End Sub
