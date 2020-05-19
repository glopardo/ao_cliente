Attribute VB_Name = "ModVisualFX"
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

Public icMode As Integer
Public UseAlphaBlending As Boolean

Public Declare Function AlphaBlend Lib "AoFX.dll" (ByVal iMode As Integer, ByVal bColorKey As Integer, ByRef sPtr As Any, ByRef dPtr As Any, ByVal iAlphaVal As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, ByVal isPitch As Integer, ByVal idPitch As Integer, ByVal iColorKey As Integer) As Integer


Sub InitBlend(surface As DirectDrawSurface7)
If UseAlphaBlending Then
    Dim ddsdtemp As DDSURFACEDESC2
          Call surface.GetSurfaceDesc(ddsdtemp)
          
          Select Case ddsdtemp.ddpfPixelFormat.lGBitMask
            Case &H3E0
                icMode = 555
            Case &H7E0
                icMode = 565
            Case Else
                MsgBox "No se pudo detectar el modo del BackBuffer ¿Esta en 16 bits de colores?"
                UseAlphaBlending = False
          End Select
End If
End Sub


Sub DDrawBlendGrhtoSurface(surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer _
, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0, Optional ByVal Blend As Byte = 150)


If Not UseAlphaBlending Then _
    Call DDrawTransGrhtoSurface(surface, Grh, X, Y, center, Animate, KillAnim)


Dim iGrhIndex As Integer, QuitarAnimacion As Boolean, rEmptyRect As RECT _
, dArray() As Byte, sArray() As Byte, source As DirectDrawSurface7, _
sourcedesc As DDSURFACEDESC2, SurfaceDesc As DDSURFACEDESC2

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                            
                            If CharList(KillAnim).FxLoopTimes > 0 Then CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes < 1 Then
                                CharList(KillAnim).FX = 0
                                Exit Sub
                            End If
                            
                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub


iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)


If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32
    End If
End If


Set source = SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum)

Call source.GetSurfaceDesc(sourcedesc)
Call surface.GetSurfaceDesc(SurfaceDesc)

surface.Lock rEmptyRect, SurfaceDesc, DDLOCK_WAIT, 0
source.Lock rEmptyRect, sourcedesc, DDLOCK_WAIT, 0

surface.GetLockedArray dArray()
source.GetLockedArray sArray()

Call AlphaBlend(icMode, 1, sArray(GrhData(iGrhIndex).sX + GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY) _
    , dArray(X + X, Y), Blend, GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
    sourcedesc.lPitch, SurfaceDesc.lPitch, 0)

source.Unlock rEmptyRect
surface.Unlock rEmptyRect

End Sub
