VERSION 5.00
Begin VB.Form Password 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Password"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   69
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   183
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Min/May"
      Height          =   195
      Left            =   2040
      TabIndex        =   37
      Top             =   795
      Width           =   630
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   36
      Left            =   480
      TabIndex        =   36
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   35
      Left            =   720
      TabIndex        =   35
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   34
      Left            =   960
      TabIndex        =   34
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   33
      Left            =   1200
      TabIndex        =   33
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   32
      Left            =   1440
      TabIndex        =   32
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   31
      Left            =   1680
      TabIndex        =   31
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   30
      Left            =   1920
      TabIndex        =   30
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   24
      Left            =   2160
      TabIndex        =   29
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   23
      Left            =   285
      TabIndex        =   28
      Top             =   0
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   22
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   29
      Left            =   840
      TabIndex        =   26
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   28
      Left            =   1080
      TabIndex        =   25
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   27
      Left            =   1320
      TabIndex        =   24
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   26
      Left            =   1560
      TabIndex        =   23
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   25
      Left            =   1800
      TabIndex        =   22
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   21
      Left            =   660
      TabIndex        =   21
      Top             =   720
      Width           =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   20
      Left            =   360
      TabIndex        =   20
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "d"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   19
      Left            =   720
      TabIndex        =   19
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "f"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   18
      Left            =   960
      TabIndex        =   18
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   17
      Left            =   1200
      TabIndex        =   17
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "h"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   16
      Left            =   1440
      TabIndex        =   16
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "j"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   15
      Left            =   1680
      TabIndex        =   15
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "k"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   14
      Left            =   1920
      TabIndex        =   14
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "l"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   13
      Left            =   2160
      TabIndex        =   13
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ñ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   12
      Left            =   2400
      TabIndex        =   12
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   540
      TabIndex        =   11
      Top             =   480
      Width           =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "w"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   390
      TabIndex        =   1
      Top             =   240
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   600
      TabIndex        =   9
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   840
      TabIndex        =   8
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   1080
      TabIndex        =   7
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   1320
      TabIndex        =   6
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   1560
      TabIndex        =   5
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   1800
      TabIndex        =   4
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "p"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "q"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   240
   End
End
Attribute VB_Name = "Password"
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

Private Sub Form_LostFocus()

Unload Me

End Sub

Private Sub Label1_Click(Index As Integer)

frmConnect.txtPass = frmConnect.txtPass & Label1(Index).Caption

End Sub
Private Sub Label2_Click()
Dim i As Integer

If Label1(0).Caption = "Q" Then
    For i = 0 To 33
        Label1(i).Caption = LCase(Label1(i).Caption)
    Next
Else
    For i = 0 To 33
        Label1(i).Caption = UCase$(Label1(i).Caption)
    Next
End If

End Sub
