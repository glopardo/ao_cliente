Attribute VB_Name = "modInventarioGrafico"
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

Public Const XCantItems = 5


Public ActualizarInv As Boolean
Public OffsetDelInv As Integer
Public ItemElegido As Integer
Public mx As Integer
Public my As Integer

Private bStaticInit  As Boolean
Private r1           As RECT, r2 As RECT, auxr As RECT
Private rBox         As RECT
Private rBoxFrame(2) As RECT
Private iFrameMod    As Integer
Sub ActualizarOtherInventory(Slot As Integer)

If OtherInventory(Slot).OBJIndex = 0 Then
    frmComerciar.List1(0).List(Slot - 1) = "Nada"
Else
    frmComerciar.List1(0).List(Slot - 1) = OtherInventory(Slot).Name
End If

If frmComerciar.List1(0).ListIndex = Slot - 1 And lista = 0 Then Call ActualizarInformacionComercio(0)

End Sub
Sub ActualizarInventario(Slot As Integer)
Dim OBJIndex As Long
Dim NameSize As Byte

ActualizarInv = True

If UserInventory(Slot).Amount = 0 Then
    'frmMain.imgObjeto(Slot).ToolTipText = "Nada"
    'frmMain.lblObjCant(Slot).ToolTipText = "Nada"
    'frmMain.lblObjCant(Slot).Caption = ""
    'If ItemElegido = Slot Then frmMain.Shape1.Visible = False
    'frmMain.Inventario.ToolTipText = "Nada"
Else
    'frmMain.Inventario.ToolTipText = UserInventory(Slot).Name
    'frmMain.imgObjeto(Slot).ToolTipText = UserInventory(Slot).Name
    'frmMain.lblObjCant(Slot).ToolTipText = UserInventory(Slot).Name
    'frmMain.lblObjCant(Slot).Caption = CStr(UserInventory(Slot).Amount)
    'If ItemElegido = Slot Then frmMain.Shape1.Visible = True
End If

'If UserInventory(Slot).GrhIndex > 0 Then
    'frmMain.imgObjeto(Slot).Picture = LoadPicture(DirGraficos & GrhData(UserInventory(Slot).GrhIndex).FileNum & ".bmp")
'Else
    'frmMain.imgObjeto(Slot).Picture = LoadPicture()
'End If

'If UserInventory(Slot).Equipped > 0 Then
'    frmMain.Label2(Slot).Visible = True
'Else
'    frmMain.Label2(Slot).Visible = False
'End If

If frmComerciar.Visible Then
    If UserInventory(Slot).Amount = 0 Then
        frmComerciar.List1(1).List(Slot - 1) = "Nada"
     Else
        frmComerciar.List1(1).List(Slot - 1) = UserInventory(Slot).Name
    End If
    If frmComerciar.List1(1).ListIndex = Slot - 1 And lista = 1 Then Call ActualizarInformacionComercio(1)
End If

End Sub
