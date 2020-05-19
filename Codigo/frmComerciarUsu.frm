VERSION 5.00
Begin VB.Form frmComerciarUsu 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   418
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCant 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   5040
      TabIndex        =   7
      Text            =   "0"
      Top             =   5175
      Width           =   1215
   End
   Begin VB.OptionButton optQue 
      BackColor       =   &H80000007&
      Height          =   195
      Index           =   0
      Left            =   4017
      TabIndex        =   6
      Top             =   960
      Value           =   -1  'True
      Width           =   195
   End
   Begin VB.OptionButton optQue 
      BackColor       =   &H80000007&
      Height          =   195
      Index           =   1
      Left            =   5310
      TabIndex        =   5
      Top             =   960
      Width           =   195
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFFFF&
      Height          =   3345
      Left            =   720
      TabIndex        =   3
      Top             =   1560
      Width           =   2850
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
      ForeColor       =   &H00FFFFFF&
      Height          =   3540
      Left            =   3960
      TabIndex        =   2
      Top             =   1320
      Width           =   2490
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   2610
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   0
      Top             =   720
      Width           =   540
   End
   Begin VB.Image cmdAceptar 
      Height          =   375
      Left            =   720
      MouseIcon       =   "frmComerciarUsu.frx":0000
      MousePointer    =   99  'Custom
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Image cmdRechazar 
      Height          =   375
      Left            =   2160
      MouseIcon       =   "frmComerciarUsu.frx":030A
      MousePointer    =   99  'Custom
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Image cmdOfrecer 
      Height          =   375
      Left            =   4560
      MouseIcon       =   "frmComerciarUsu.frx":0614
      MousePointer    =   99  'Custom
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   4020
      TabIndex        =   8
      Top             =   4560
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   5175
      Width           =   975
   End
   Begin VB.Image command2 
      Height          =   255
      Left            =   0
      MouseIcon       =   "frmComerciarUsu.frx":091E
      MousePointer    =   99  'Custom
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label lblEstadoResp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Esperando respuesta..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   2490
   End
End
Attribute VB_Name = "frmComerciarUsu"
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



Private Sub cmdAceptar_Click()
Call SendData("COMUSUOK")
End Sub

Private Sub cmdOfrecer_Click()

If optQue(0).value = True Then
    If List1.ListIndex < 0 Then Exit Sub
    If List1.ItemData(List1.ListIndex) <= 0 Then Exit Sub
    


ElseIf optQue(1).value = True Then



End If

If optQue(0).value = True Then
    Call SendData("OFRECER" & List1.ListIndex + 1 & "," & Trim(Val(txtCant.Text)))
ElseIf optQue(1).value = True Then
    Call SendData("OFRECER" & FLAGORO & "," & Trim(Val(txtCant.Text)))
Else
    Exit Sub
End If

lblEstadoResp.Visible = True

End Sub

Private Sub cmdRechazar_Click()
Call SendData("COMUSUNO")
End Sub

Private Sub Command2_Click()
Call SendData("FINCOMUSU")

End Sub

Private Sub Form_Deactivate()

Me.SetFocus
Picture1.SetFocus

End Sub
Private Sub Form_Load()

lblEstadoResp.Visible = False
Me.Picture = LoadPicture(DirGraficos & "ComerciarUsu.gif")

End Sub
Private Sub Form_LostFocus()

Me.SetFocus
Picture1.SetFocus

End Sub
Private Sub list1_Click()

DibujaGrh UserInventory(List1.ListIndex + 1).GrhIndex

End Sub
Public Sub DibujaGrh(Grh As Integer)
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.Bottom = 32

Call DrawGrhtoHdc(Picture1.hwnd, Picture1.Hdc, Grh, SR, DR)

End Sub
Private Sub List2_Click()

If List2.ListIndex >= 0 Then
    DibujaGrh OtroInventario(List2.ListIndex + 1).GrhIndex
    Label3.Caption = List2.ItemData(List2.ListIndex)
    cmdAceptar.Enabled = True
    cmdRechazar.Enabled = True
Else
    cmdAceptar.Enabled = False
    cmdRechazar.Enabled = False
End If

End Sub
Private Sub optQue_Click(Index As Integer)

Select Case Index
    Case 0
        List1.Enabled = True
    Case 1
        List1.Enabled = False
End Select

End Sub
Private Sub txtCant_KeyDown(KeyCode As Integer, Shift As Integer)

If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = vbKeyBack Or _
        KeyCode = vbKeyDelete Or (KeyCode >= 37 And KeyCode <= 40)) Then
    
    KeyCode = 0
End If

End Sub
Private Sub txtCant_KeyPress(KeyAscii As Integer)

If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or _
        KeyAscii = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
    
    KeyAscii = 0
End If

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving = False And Button = vbLeftButton Then
    DX = X
    dy = Y
    bmoving = True
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving And ((X <> DX) Or (Y <> dy)) Then Call Move(Left + (X - DX), Top + (Y - dy))

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub
