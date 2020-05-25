Attribute VB_Name = "modGeneralCharFunctions"
'Parra: Este modulo contiene funciones generales relacionadas con el Char o con el movimiento del mismo _
                que antes estaban en el TileEngine y en General pero que prefiero pasarlas aqui para tener el TileEngine más limpio
                
Option Explicit

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal x As Integer, ByVal y As Integer, ByVal arma As Integer, ByVal escudo As Integer, ByVal casco As Integer)
On Error Resume Next


If CharIndex > LastChar Then LastChar = CharIndex

If arma = 0 Then arma = 2
If escudo = 0 Then escudo = 2
If casco = 0 Then casco = 2

CharList(CharIndex).Head = HeadData(Head)

CharList(CharIndex).Body = BodyData(Body)

If Body > 83 And Body < 88 Then
    CharList(CharIndex).Navegando = 1
Else: CharList(CharIndex).Navegando = 0
End If

CharList(CharIndex).arma = WeaponAnimData(arma)
    
CharList(CharIndex).escudo = ShieldAnimData(escudo)
CharList(CharIndex).casco = CascoAnimData(casco)

CharList(CharIndex).Heading = Heading


CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.x = 0
CharList(CharIndex).MoveOffset.y = 0


CharList(CharIndex).POS.x = x
CharList(CharIndex).POS.y = y


CharList(CharIndex).Active = 1


MapData(x, y).CharIndex = CharIndex

End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)

CharList(CharIndex).Active = 0
CharList(CharIndex).Criminal = 0
CharList(CharIndex).FX = 0
CharList(CharIndex).FxLoopTimes = 0
CharList(CharIndex).invisible = False
CharList(CharIndex).Moving = 0
CharList(CharIndex).muerto = False
CharList(CharIndex).Nombre = ""
CharList(CharIndex).pie = False
CharList(CharIndex).POS.x = 0
CharList(CharIndex).POS.y = 0
CharList(CharIndex).UsandoArma = False

End Sub
Function NextOpenChar()
Dim loopc As Integer

loopc = 1

Do While CharList(loopc).Active
    loopc = loopc + 1
Loop

NextOpenChar = loopc

End Function
Sub EraseChar(ByVal CharIndex As Integer)
On Error Resume Next

CharList(CharIndex).Active = 0


If CharIndex = LastChar Then
    Do Until CharList(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If


MapData(CharList(CharIndex).POS.x, CharList(CharIndex).POS.y).CharIndex = 0

Call ResetCharInfo(CharIndex)

End Sub

Public Sub DoFogataFx()
If FX = 0 Then
    If bFogata Then
        bFogata = HayFogata()
        If Not bFogata Then Audio.StopWave
    Else
        bFogata = HayFogata()
        If bFogata Then Audio.PlayWave "fuego.wav", 0, 0, Enabled
    End If
End If
End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean

Dim x As Integer, y As Integer

For y = UserPos.y - MinYBorder + 1 To UserPos.y + MinYBorder - 1
  For x = UserPos.x - MinXBorder + 1 To UserPos.x + MinXBorder - 1
            
            If MapData(x, y).CharIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
  Next x
Next y

EstaPCarea = False

End Function
Public Function TickON(Cual As Integer, Cont As Integer) As Boolean
Static TickCount(200) As Integer
If Cont = 999 Then Exit Function
TickCount(Cual) = TickCount(Cual) + 1
If TickCount(Cual) < Cont Then
    TickON = False
Else
    TickCount(Cual) = 0
    TickON = True
End If
End Function
Sub DoPasosFx(ByVal CharIndex As Integer)
Static pie As Boolean

If CharList(CharIndex).Navegando = 0 Then
    If UserMontando And EstaPCarea(CharIndex) And CharIndex = UserCharIndex Then
        If TickON(0, 4) Then Call Audio.PlayWave(SND_MONTANDO)
    Else
        If CharList(CharIndex).Criminal = 1 Then Exit Sub
        If Not CharList(CharIndex).muerto And EstaPCarea(CharIndex) Then
            CharList(CharIndex).pie = Not CharList(CharIndex).pie
            If CharList(CharIndex).pie Then
                Call Audio.PlayWave(SND_PASOS1)
            Else
                Call Audio.PlayWave(SND_PASOS2)
            End If
        End If
    End If
Else: Call Audio.PlayWave(SND_NAVEGANDO)
End If

End Sub



Sub MoveMe(Direction As Byte)

If CONGELADO Then Exit Sub

If Cartel Then Cartel = False

If ProxLegalPos(Direction) And Not UserMeditar And Not UserParalizado Then
    If TiempoTranscurrido(LastPaso) >= IntervaloPaso Then
        Call SendData("M" & Direction)
        Call DibujarMiniMapa
        LastPaso = Timer
        If Not UserDescansar Then
            Call EliminarChars(Direction)
            Call MoveCharByHead(UserCharIndex, Direction)
            Call MoveScreen(Direction)
            Call DoFogataFx
        End If
    End If
ElseIf CharList(UserCharIndex).Heading <> Direction Then Call SendData("CHEA" & Direction)
End If

frmMain.mapa.Caption = NombreDelMapaActual & " [" & UserMap & " - " & UserPos.x & " - " & UserPos.y & "]"

End Sub
Function ProxLegalPos(Direction As Byte) As Boolean

Select Case Direction
    Case NORTH
        ProxLegalPos = LegalPos(UserPos.x, UserPos.y - 1)
    Case SOUTH
        ProxLegalPos = LegalPos(UserPos.x, UserPos.y + 1)
    Case WEST
        ProxLegalPos = LegalPos(UserPos.x - 1, UserPos.y)
    Case EAST
        ProxLegalPos = LegalPos(UserPos.x + 1, UserPos.y)
End Select

End Function

Sub MoveScreen(Heading As Byte)

Dim x As Integer
Dim y As Integer
Dim tX As Integer
Dim tY As Integer

Select Case Heading

    Case NORTH
        y = -1

    Case EAST
        x = 1

    Case SOUTH
        y = 1
    
    Case WEST
        x = -1
        
End Select


tX = UserPos.x + x
tY = UserPos.y + y


If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
    Exit Sub
Else
    AddtoUserPos.x = x
    UserPos.x = tX
    AddtoUserPos.y = y
    UserPos.y = tY
    UserMoving = 1
    bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or MapData(UserPos.x, UserPos.y).Trigger = 2 Or MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
End If

End Sub
Function HayFogata() As Boolean
Dim j As Integer, k As Integer
For j = UserPos.x - 8 To UserPos.x + 8
    For k = UserPos.y - 6 To UserPos.y + 6
        If InMapBounds(j, k) Then
            If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    HayFogata = True
                    Exit Function
            End If
        End If
    Next k
Next j
End Function
Sub RefreshAllChars()
Dim loopc As Integer

For loopc = 1 To LastChar
    If CharList(loopc).Active = 1 Then
        MapData(CharList(loopc).POS.x, CharList(loopc).POS.y).CharIndex = loopc
    End If
Next loopc

End Sub
Function LegalPos(x As Integer, y As Integer) As Boolean

    If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        LegalPos = False
        Exit Function
    End If

    If MapData(x, y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
    If MapData(x, y).CharIndex > 0 Then
        LegalPos = False
        Exit Function
    End If
   
    If Not UserNavegando Then
        If HayAgua(x, y) Then
            LegalPos = False
            Exit Function
        End If
    Else
        If Not HayAgua(x, y) Then
            LegalPos = False
            Exit Function
        End If
    End If
    
LegalPos = True

End Function
Function InMapBounds(x As Integer, y As Integer) As Boolean
    If x < XMinMapSize Or x > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
        InMapBounds = False
        Exit Function
    End If

    InMapBounds = True
End Function
Function HayAgua(x As Integer, y As Integer) As Boolean

If MapData(x, y).Graphic(1).GrhIndex >= 1505 And _
   MapData(x, y).Graphic(1).GrhIndex <= 1520 And _
   MapData(x, y).Graphic(2).GrhIndex = 0 Then
            HayAgua = True
Else
            HayAgua = False
End If

End Function
Sub EliminarChars(Direction As Byte)
Dim x(2) As Integer
Dim y(2) As Integer

Select Case Direction
    Case NORTH, SOUTH
        x(1) = UserPos.x - MinXBorder - 2
        x(2) = UserPos.x + MinXBorder + 2
    Case EAST, WEST
        y(1) = UserPos.y - MinYBorder - 2
        y(2) = UserPos.y + MinYBorder + 2
End Select

Select Case Direction
    Case NORTH
        y(1) = UserPos.y - MinYBorder - 3
        If y(1) < 1 Then y(1) = 1
        y(2) = y(1)
    Case EAST
        x(1) = UserPos.x + MinXBorder + 3
        If x(1) > 99 Then x(1) = 99
        x(2) = x(1)
    Case SOUTH
        y(1) = UserPos.y + MinYBorder + 3
        If y(1) > 99 Then y(1) = 99
        y(2) = y(1)
    Case WEST
        x(1) = UserPos.x - MinXBorder - 3
        If x(1) < 1 Then x(1) = 1
        x(2) = x(1)
End Select

For y(0) = y(1) To y(2)
    For x(0) = x(1) To x(2)
        If x(0) > 6 And x(0) < 95 And y(0) > 6 And y(0) < 95 Then
            If MapData(x(0), y(0)).CharIndex > 0 Then
                CharList(MapData(x(0), y(0)).CharIndex).POS.x = 0
                CharList(MapData(x(0), y(0)).CharIndex).POS.y = 0
                MapData(x(0), y(0)).CharIndex = 0
            End If
        End If
    Next
Next

End Sub

Sub MoveCharByPosAndHead(CharIndex As Integer, nX As Integer, nY As Integer, nheading As Byte)
 
On Error Resume Next
 
Dim x As Integer
Dim y As Integer
Dim addX As Integer
Dim addY As Integer
 
 
 
x = CharList(CharIndex).POS.x
y = CharList(CharIndex).POS.y
 
MapData(x, y).CharIndex = 0
 
addX = nX - x
addY = nY - y
 
MapData(nX, nY).CharIndex = CharIndex
 
CharList(CharIndex).POS.x = nX
CharList(CharIndex).POS.y = nY
 
CharList(CharIndex).MoveOffset.x = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.y = -1 * (TilePixelHeight * addY)
 
CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nheading
 
CharList(CharIndex).scrollDirectionX = Sgn(addX)
CharList(CharIndex).scrollDirectionY = Sgn(addY)
 
'.MoveOffset.X = -1 * (32 * addX)
'.MoveOffset.Y = -1 * (32 * addY)
 
'.Moving = 1
'.Heading = nheading
 
'.scrollDirectionX = addX
'.scrollDirectionY = addY
 
 
End Sub
Sub MoveCharByPos(CharIndex As Integer, nX As Integer, nY As Integer)
On Error Resume Next
 
Dim x As Integer
Dim y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nheading As Byte
 
x = CharList(CharIndex).POS.x
y = CharList(CharIndex).POS.y
 
'MapData(X, y).CharIndex = 0
 
addX = nX - x
addY = nY - y
 
If Sgn(addX) = -1 Then nheading = WEST
If Sgn(addX) = 1 Then nheading = EAST
 
If Sgn(addY) = -1 Then nheading = NORTH
If Sgn(addY) = 1 Then nheading = SOUTH
 
'MapData(nX, nY).CharIndex = CharIndex
 
'CharList(CharIndex).POS.X = nX
'CharList(CharIndex).POS.y = nY
 
'CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
'CharList(CharIndex).MoveOffset.y = -1 * (TilePixelHeight * addY)
 
'CharList(CharIndex).Moving = 1
'CharList(CharIndex).Heading = nheading
 
'CharList(CharIndex).scrollDirectionX = Sgn(addX)
'CharList(CharIndex).scrollDirectionY = Sgn(addY)
 
MoveCharByHead CharIndex, nheading
 
 
 
End Sub
Sub MoveCharByPosConHeading(CharIndex As Integer, nX As Integer, nY As Integer, nheading As Byte)
On Error Resume Next
 
If InMapBounds(CharList(CharIndex).POS.x, CharList(CharIndex).POS.y) Then MapData(CharList(CharIndex).POS.x, CharList(CharIndex).POS.y).CharIndex = 0
 
MapData(nX, nY).CharIndex = CharIndex
 
CharList(CharIndex).POS.x = nX
CharList(CharIndex).POS.y = nY
 
CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.x = 0
CharList(CharIndex).MoveOffset.y = 0
 
CharList(CharIndex).Heading = nheading
 
End Sub
 
Sub MoveCharByHead(CharIndex As Integer, nheading As Byte)
 
Dim addX As Integer
Dim addY As Integer
Dim x As Integer
Dim y As Integer
Dim nX As Integer
Dim nY As Integer
 
With CharList(CharIndex)
x = .POS.x
y = .POS.y
 
 
Select Case nheading
 
    Case NORTH
        addY = -1
 
    Case EAST
        addX = 1
 
    Case SOUTH
        addY = 1
   
    Case WEST
        addX = -1
       
End Select
 
nX = x + addX
nY = y + addY
 
MapData(nX, nY).CharIndex = CharIndex
.POS.x = nX
.POS.y = nY
MapData(x, y).CharIndex = 0
 
.MoveOffset.x = -1 * (32 * addX)
.MoveOffset.y = -1 * (32 * addY)
 
.Moving = 1
.Heading = nheading
 
.scrollDirectionX = addX
.scrollDirectionY = addY
 
If UserEstado <> True Then Call DoPasosFx(CharIndex)
 
 
End With
 
End Sub
