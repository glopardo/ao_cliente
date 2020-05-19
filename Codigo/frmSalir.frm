VERSION 5.00
Begin VB.Form frmSalir 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image cancelar 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1920
      MouseIcon       =   "frmSalir.frx":0000
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Image aceptar 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   720
      MouseIcon       =   "frmSalir.frx":030A
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "frmSalir"
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


Private Sub Aceptar_Click()

Call SendData("/SALIR")
Unload Me
Unload frmMain

End Sub

Private Sub Cancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "Salir.gif")


End Sub
