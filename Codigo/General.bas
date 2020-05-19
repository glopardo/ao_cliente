Attribute VB_Name = "modGeneral"
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

Public Audio As New clsAudio

'Sound constants
Public Const SND_CLICK = "click.Wav"
Public Const SND_MONTANDO = "23.Wav"
Public Const SND_PASOS1 = "23.Wav"
Public Const SND_PASOS2 = "24.Wav"
Public Const SND_NAVEGANDO = "50.wav"
Public Const SND_OVER = "click2.Wav"
Public Const SND_DICE = "cupdice.Wav"
Public Const MIdi_Inicio = 6


Public CartelOcultarse As Byte
Public CartelMenosCansado As Byte
Public CartelVestirse As Byte
Public CartelNoHayNada As Byte
Public CartelRecuMana As Byte
Public CartelSanado As Byte
Public atacar As Integer
Public IsClan As Byte
Public NoRes As Boolean
Public ResOriginal As Boolean
Public Desplazar As Boolean
Public vigilar As Boolean


Public RG(1 To 5, 1 To 3) As Byte

Public bO As Integer
Public bK As Long
Public bRK As Long
Public banners As String

Public bInvMod     As Boolean

Public bFogata As Boolean

Public bLluvia() As Byte

Type Recompensa
    Name As String
    Descripcion As String
End Type

Public Recompensas(1 To 60, 1 To 3, 1 To 2) As Recompensa
Public Sub EstablecerRecompensas()

Recompensas(MINERO, 1, 1).Name = "Fortaleza del Trabajador"
Recompensas(MINERO, 1, 1).Descripcion = "Aumenta la vida en 120 puntos."

Recompensas(MINERO, 1, 2).Name = "Suerte de Novato"
Recompensas(MINERO, 1, 2).Descripcion = "Al morir hay 20% de probabilidad de no perder los minerales."

Recompensas(MINERO, 2, 1).Name = "Destrucción Mágica"
Recompensas(MINERO, 2, 1).Descripcion = "Inmunidad al paralisis lanzado por otros usuarios."

Recompensas(MINERO, 2, 2).Name = "Pica Fuerte"
Recompensas(MINERO, 2, 2).Descripcion = "Permite minar 20% más cantidad de hierro y la plata."

Recompensas(MINERO, 3, 1).Name = "Gremio del Trabajador"
Recompensas(MINERO, 3, 1).Descripcion = "Permite minar 20% más cantidad de oro."

Recompensas(MINERO, 3, 2).Name = "Pico de la Suerte"
Recompensas(MINERO, 3, 2).Descripcion = "Al morir hay 30% de probabilidad de que no perder los minerales (acumulativo con Suerte de Novato.)"


Recompensas(HERRERO, 1, 1).Name = "Yunque Rojizo"
Recompensas(HERRERO, 1, 1).Descripcion = "25% de probabilidad de gastar la mitad de lingotes en la creación de objetos (Solo aplicable a armas y armaduras)."

Recompensas(HERRERO, 1, 2).Name = "Maestro de la Forja"
Recompensas(HERRERO, 1, 2).Descripcion = "Reduce los costos de cascos y escudos a un 50%."

Recompensas(HERRERO, 2, 1).Name = "Experto en Filos"
Recompensas(HERRERO, 2, 1).Descripcion = "Permite crear las mejores armas (Espada Neithan, Espada Neithan + 1, Espada de Plata + 1 y Daga Infernal)."

Recompensas(HERRERO, 2, 2).Name = "Experto en Corazas"
Recompensas(HERRERO, 2, 2).Descripcion = "Permite crear las mejores armaduras (Armaduras de las Tinieblas, Armadura Legendaria y Armaduras del Dragón)."

Recompensas(HERRERO, 3, 1).Name = "Fundir Metal"
Recompensas(HERRERO, 3, 1).Descripcion = "Reduce a un 50% la cantidad de lingotes utilizados en fabricación de Armas y Armaduras (acumulable con Yunque Rojizo)."

Recompensas(HERRERO, 3, 2).Name = "Trabajo en Serie"
Recompensas(HERRERO, 3, 2).Descripcion = "10% de probabilidad de crear el doble de objetos de los asignados con la misma cantidad de lingotes."


Recompensas(TALADOR, 1, 1).Name = "Músculos Fornidos"
Recompensas(TALADOR, 1, 1).Descripcion = "Permite talar 20% más cantidad de madera."

Recompensas(TALADOR, 1, 2).Name = "Tiempos de Calma"
Recompensas(TALADOR, 1, 2).Descripcion = "Evita tener hambre y sed."


Recompensas(CARPINTERO, 1, 1).Name = "Experto en Arcos"
Recompensas(CARPINTERO, 1, 1).Descripcion = "Permite la creación de los mejores arcos (Élfico y de las Tinieblas)."

Recompensas(CARPINTERO, 1, 2).Name = "Experto de Varas"
Recompensas(CARPINTERO, 1, 2).Descripcion = "Permite la creación de las mejores varas (Engarzadas)."

Recompensas(CARPINTERO, 2, 1).Name = "Fila de Leña"
Recompensas(CARPINTERO, 2, 1).Descripcion = "Aumenta la creación de flechas a 20 por vez."

Recompensas(CARPINTERO, 2, 2).Name = "Espíritu de Navegante"
Recompensas(CARPINTERO, 2, 2).Descripcion = "Reduce en un 20% el coste de madera de las barcas."


Recompensas(PESCADOR, 1, 1).Name = "Favor de los Dioses"
Recompensas(PESCADOR, 1, 1).Descripcion = "Pescar 20% más cantidad de pescados."

Recompensas(PESCADOR, 1, 2).Name = "Pesca en Alta Mar"
Recompensas(PESCADOR, 1, 2).Descripcion = "Al pescar en barca hay 10% de probabilidad de obtener pescados más caros."


Recompensas(MAGO, 1, 1).Name = "Pociones de Espíritu"
Recompensas(MAGO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(MAGO, 1, 2).Name = "Pociones de Vida"
Recompensas(MAGO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(MAGO, 2, 1).Name = "Vitalidad"
Recompensas(MAGO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(MAGO, 2, 2).Name = "Fortaleza Mental"
Recompensas(MAGO, 2, 2).Descripcion = "Libera el limite de mana máximo."

Recompensas(MAGO, 3, 1).Name = "Furia del Relámpago"
Recompensas(MAGO, 3, 1).Descripcion = "Aumenta el daño base máximo de la Descarga Eléctrica en 10 puntos."

Recompensas(MAGO, 3, 2).Name = "Destrucción"
Recompensas(MAGO, 3, 2).Descripcion = "Aumenta el daño base mínimo del Apocalipsis en 10 puntos."


Recompensas(NIGROMANTE, 1, 1).Name = "Pociones de Espíritu"
Recompensas(NIGROMANTE, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(NIGROMANTE, 1, 2).Name = "Pociones de Vida"
Recompensas(NIGROMANTE, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(NIGROMANTE, 2, 1).Name = "Vida del Invocador"
Recompensas(NIGROMANTE, 2, 1).Descripcion = "Aumenta la vida en 15 puntos."

Recompensas(NIGROMANTE, 2, 2).Name = "Alma del Invocador"
Recompensas(NIGROMANTE, 2, 2).Descripcion = "Aumenta el mana en 40 puntos."

Recompensas(NIGROMANTE, 3, 1).Name = "Semillas de las Almas"
Recompensas(NIGROMANTE, 3, 1).Descripcion = "Aumenta el daño base mínimo de la magia en 10 puntos."

Recompensas(NIGROMANTE, 3, 2).Name = "Bloqueo de las Almas"
Recompensas(NIGROMANTE, 3, 2).Descripcion = "Aumenta la evasión en un 5%."


Recompensas(PALADIN, 1, 1).Name = "Pociones de Espíritu"
Recompensas(PALADIN, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(PALADIN, 1, 2).Name = "Pociones de Vida"
Recompensas(PALADIN, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(PALADIN, 2, 1).Name = "Aura de Vitalidad"
Recompensas(PALADIN, 2, 1).Descripcion = "Aumenta la vida en 5 puntos y el mana en 10 puntos."

Recompensas(PALADIN, 2, 2).Name = "Aura de Espíritu"
Recompensas(PALADIN, 2, 2).Descripcion = "Aumenta el mana en 30 puntos."

Recompensas(PALADIN, 3, 1).Name = "Gracia Divina"
Recompensas(PALADIN, 3, 1).Descripcion = "Reduce el coste de mana de Remover Paralisis a 250 puntos."

Recompensas(PALADIN, 3, 2).Name = "Favor de los Enanos"
Recompensas(PALADIN, 3, 2).Descripcion = "Aumenta en 5% la posibilidad de golpear al enemigo con armas cuerpo a cuerpo."


Recompensas(CLERIGO, 1, 1).Name = "Pociones de Espíritu"
Recompensas(CLERIGO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(CLERIGO, 1, 2).Name = "Pociones de Vida"
Recompensas(CLERIGO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(CLERIGO, 2, 1).Name = "Signo Vital"
Recompensas(CLERIGO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(CLERIGO, 2, 2).Name = "Espíritu de Sacerdote"
Recompensas(CLERIGO, 2, 2).Descripcion = "Aumenta el mana en 50 puntos."

Recompensas(CLERIGO, 3, 1).Name = "Sacerdote Experto"
Recompensas(CLERIGO, 3, 1).Descripcion = "Aumenta la cura base de Curar Heridas Graves en 20 puntos."

Recompensas(CLERIGO, 3, 2).Name = "Alzamientos de Almas"
Recompensas(CLERIGO, 3, 2).Descripcion = "El hechizo de Resucitar cura a las personas con su mana, energía, hambre y sed llenas y cuesta 1.100 de mana."


Recompensas(BARDO, 1, 1).Name = "Pociones de Espíritu"
Recompensas(BARDO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(BARDO, 1, 2).Name = "Pociones de Vida"
Recompensas(BARDO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(BARDO, 2, 1).Name = "Melodía Vital"
Recompensas(BARDO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(BARDO, 2, 2).Name = "Melodía de la Meditación"
Recompensas(BARDO, 2, 2).Descripcion = "Aumenta el mana en 50 puntos."

Recompensas(BARDO, 3, 1).Name = "Concentración"
Recompensas(BARDO, 3, 1).Descripcion = "Aumenta la probabilidad de Apuñalar a un 20% (con 100 skill)."

Recompensas(BARDO, 3, 2).Name = "Melodía Caótica"
Recompensas(BARDO, 3, 2).Descripcion = "Aumenta el daño base del Apocalipsis y la Descarga Electrica en 5 puntos."


Recompensas(DRUIDA, 1, 1).Name = "Pociones de Espíritu"
Recompensas(DRUIDA, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(DRUIDA, 1, 2).Name = "Pociones de Vida"
Recompensas(DRUIDA, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(DRUIDA, 2, 1).Name = "Grifo de la Vida"
Recompensas(DRUIDA, 2, 1).Descripcion = "Aumenta la vida en 15 puntos."

Recompensas(DRUIDA, 2, 2).Name = "Poder del Alma"
Recompensas(DRUIDA, 2, 2).Descripcion = "Aumenta el mana en 40 puntos."

Recompensas(DRUIDA, 3, 1).Name = "Raíces de la Naturaleza"
Recompensas(DRUIDA, 3, 1).Descripcion = "Reduce el coste de mana de Inmovilizar a 250 puntos."

Recompensas(DRUIDA, 3, 2).Name = "Fortaleza Natural"
Recompensas(DRUIDA, 3, 2).Descripcion = "Aumenta la vida de los elementales invocados en 75 puntos."


Recompensas(ASESINO, 1, 1).Name = "Pociones de Espíritu"
Recompensas(ASESINO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(ASESINO, 1, 2).Name = "Pociones de Vida"
Recompensas(ASESINO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(ASESINO, 2, 1).Name = "Sombra de Vida"
Recompensas(ASESINO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(ASESINO, 2, 2).Name = "Sombra Mágica"
Recompensas(ASESINO, 2, 2).Descripcion = "Aumenta el mana en 30 puntos."

Recompensas(ASESINO, 3, 1).Name = "Daga Mortal"
Recompensas(ASESINO, 3, 1).Descripcion = "Aumenta el daño de Apuñalar a un 70% más que el golpe."

Recompensas(ASESINO, 3, 2).Name = "Punteria mortal"
Recompensas(ASESINO, 3, 2).Descripcion = "Las chances de apuñalar suben a 25% (Con 100 skills)."


Recompensas(CAZADOR, 1, 1).Name = "Pociones de Espíritu"
Recompensas(CAZADOR, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(CAZADOR, 1, 2).Name = "Pociones de Vida"
Recompensas(CAZADOR, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(CAZADOR, 2, 1).Name = "Fortaleza del Oso"
Recompensas(CAZADOR, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(CAZADOR, 2, 2).Name = "Fortaleza del Leviatán"
Recompensas(CAZADOR, 2, 2).Descripcion = "Aumenta el mana en 50 puntos."

Recompensas(CAZADOR, 3, 1).Name = "Precisión"
Recompensas(CAZADOR, 3, 1).Descripcion = "Aumenta la puntería con arco en un 10%."

Recompensas(CAZADOR, 3, 2).Name = "Tiro Preciso"
Recompensas(CAZADOR, 3, 2).Descripcion = "Las flechas que golpeen la cabeza ignoran la defensa del casco."


Recompensas(ARQUERO, 1, 1).Name = "Flechas Mortales"
Recompensas(ARQUERO, 1, 1).Descripcion = "1.500 flechas que caen al morir."

Recompensas(ARQUERO, 1, 2).Name = "Pociones de Vida"
Recompensas(ARQUERO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(ARQUERO, 2, 1).Name = "Vitalidad Élfica"
Recompensas(ARQUERO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(ARQUERO, 2, 2).Name = "Paso Élfico"
Recompensas(ARQUERO, 2, 2).Descripcion = "Aumenta la evasión en un 5%."

Recompensas(ARQUERO, 3, 1).Name = "Ojo del Águila"
Recompensas(ARQUERO, 3, 1).Descripcion = "Aumenta la puntería con arco en un 5%."

Recompensas(ARQUERO, 3, 2).Name = "Disparo Élfico"
Recompensas(ARQUERO, 3, 2).Descripcion = "Aumenta el daño base mínimo de las flechas en 5 puntos y el máximo en 3 puntos."


Recompensas(GUERRERO, 1, 1).Name = "Pociones de Poder"
Recompensas(GUERRERO, 1, 1).Descripcion = "80 pociones verdes y 100 amarillas que no caen al morir."

Recompensas(GUERRERO, 1, 2).Name = "Pociones de Vida"
Recompensas(GUERRERO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(GUERRERO, 2, 1).Name = "Vida del Mamut"
Recompensas(GUERRERO, 2, 1).Descripcion = "Aumenta la vida en 5 puntos."

Recompensas(GUERRERO, 2, 2).Name = "Piel de Piedra"
Recompensas(GUERRERO, 2, 2).Descripcion = "Aumenta la defensa permanentemente en 2 puntos."

Recompensas(GUERRERO, 3, 1).Name = "Cuerda Tensa"
Recompensas(GUERRERO, 3, 1).Descripcion = "Aumenta la puntería con arco en un 10%."

Recompensas(GUERRERO, 3, 2).Name = "Resistencia Mágica"
Recompensas(GUERRERO, 3, 2).Descripcion = "Reduce la duración de la parálisis de un minuto a 45 segundos."


Recompensas(PIRATA, 1, 1).Name = "Marejada Vital"
Recompensas(PIRATA, 1, 1).Descripcion = "Aumenta la vida en 20 puntos."

Recompensas(PIRATA, 1, 2).Name = "Aventurero Arriesgado"
Recompensas(PIRATA, 1, 2).Descripcion = "Permite entrar a los dungeons independientemente del nivel."

Recompensas(PIRATA, 2, 1).Name = "Riqueza"
Recompensas(PIRATA, 2, 1).Descripcion = "10% de probabilidad de no perder los objetos al morir."

Recompensas(PIRATA, 2, 2).Name = "Escamas del Dragón"
Recompensas(PIRATA, 2, 2).Descripcion = "Aumenta la vida en 40 puntos."

Recompensas(PIRATA, 3, 1).Name = "Magia Tabú"
Recompensas(PIRATA, 3, 1).Descripcion = "Inmunidad a la paralisis."

Recompensas(PIRATA, 3, 2).Name = "Cuerda de Escape"
Recompensas(PIRATA, 3, 2).Descripcion = "Permite salir del juego en solo dos segundos."


Recompensas(LADRON, 1, 1).Name = "Codicia"
Recompensas(LADRON, 1, 1).Descripcion = "Aumenta en 10% la cantidad de oro robado."

Recompensas(LADRON, 1, 2).Name = "Manos Sigilosas"
Recompensas(LADRON, 1, 2).Descripcion = "Aumenta en 5% la probabilidad de robar exitosamente."

Recompensas(LADRON, 2, 1).Name = "Pies sigilosos"
Recompensas(LADRON, 2, 1).Descripcion = "Permite moverse mientrás se está oculto."

Recompensas(LADRON, 2, 2).Name = "Ladrón Experto"
Recompensas(LADRON, 2, 2).Descripcion = "Permite el robo de objetos (10% de probabilidad)."

Recompensas(LADRON, 3, 1).Name = "Robo Lejano"
Recompensas(LADRON, 3, 1).Descripcion = "Permite robar a una distancia de hasta 4 tiles."

Recompensas(LADRON, 3, 2).Name = "Fundido de Sombra"
Recompensas(LADRON, 3, 2).Descripcion = "Aumenta en 10% la probabilidad de robar objetos."

End Sub

Public Function DirGraficos() As String
DirGraficos = App.Path & "\Graficos\"
End Function
Public Function SD(ByVal n As Integer) As Integer

Dim auxint As Integer
Dim digit As Byte
Dim suma As Integer
auxint = n

Do
    digit = (auxint Mod 10)
    suma = suma + digit
    auxint = auxint \ 10

Loop While (auxint <> 0)

SD = suma

End Function

Public Function SDM(ByVal n As Integer) As Integer

Dim auxint As Integer
Dim digit As Integer
Dim suma As Integer
auxint = n

Do
    digit = (auxint Mod 10)
    
    digit = digit - 1
    
    suma = suma + digit
    
    auxint = auxint \ 10

Loop While (auxint <> 0)

SDM = suma

End Function

Public Function Complex(ByVal n As Integer) As Integer

If n Mod 2 <> 0 Then
    Complex = n * SD(n)
Else
    Complex = n * SDM(n)
End If

End Function

Public Function ValidarLoginMSG(ByVal n As Integer) As Integer
Dim AuxInteger As Integer
Dim AuxInteger2 As Integer
AuxInteger = SD(n)
AuxInteger2 = SDM(n)
ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function
Sub AddtoRichTextBox(RichTextBox As RichTextBox, Text As String, Optional Red As Integer = -1, Optional Green As Integer, Optional Blue As Integer, Optional Bold As Boolean, Optional Italic As Boolean, Optional bCrLf As Boolean)

With RichTextBox
    If (Len(.Text)) > 4000 Then .Text = ""
    .SelStart = Len(RichTextBox.Text)
    .SelLength = 0

    .SelBold = IIf(Bold, True, False)
    .SelItalic = IIf(Italic, True, False)
    
    If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)

    .SelText = IIf(bCrLf, Text, Text & vbCrLf)
    
    RichTextBox.Refresh
End With

End Sub
Sub AddtoTextBox(TextBox As TextBox, Text As String)

TextBox.SelStart = Len(TextBox.Text)
TextBox.SelLength = 0

TextBox.SelText = Chr(13) & Chr(10) & Text

End Sub
Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function



Function CheckUserData(checkemail As Boolean) As Boolean

Dim loopc As Integer
Dim CharAscii As Integer

If checkemail Then
 If UserEmail = "" Then
    MsgBox ("Direccion de email invalida")
    Exit Function
 End If
End If

If UserPassword = "" Then
    MsgBox "Ingrese la contraseña de su personaje.", vbInformation, "Password"
    Exit Function
End If

For loopc = 1 To Len(UserPassword)
    CharAscii = Asc(mid$(UserPassword, loopc, 1))
    If LegalCharacter(CharAscii) = False Then
        MsgBox "El password es inválido." & vbCrLf & vbCrLf & "Volvé a intentarlo otra vez." & vbCrLf & "Si el password es ese, verifica el estado del BloqMayús.", vbExclamation, "Password inválido"
        Exit Function
    End If
Next loopc

If UserName = "" Then
    MsgBox "Tenés que ingresar el Nombre de tu Personaje para poder Jugar.", vbExclamation, "Nombre inválido"
    Exit Function
End If

If Len(UserName) > 20 Then
    MsgBox ("El Nombre de tu Personaje debe tener menos de 20 letras.")
    Exit Function
End If

For loopc = 1 To Len(UserName)

    CharAscii = Asc(mid$(UserName, loopc, 1))
    If LegalCharacter(CharAscii) = False Then
        MsgBox "El Nombre del Personaje ingresado es inválido." & vbCrLf & vbCrLf & "Verifica que no halla errores en el tipeo del Nombre de tu Personaje.", vbExclamation, "Carácteres inválidos"
        Exit Function
    End If
    
Next loopc


CheckUserData = True

End Function
Sub UnloadAllForms()
On Error Resume Next
Dim mifrm As Form

For Each mifrm In Forms
    Unload mifrm
Next

End Sub

Function LegalCharacter(KeyAscii As Integer) As Boolean

If KeyAscii = 8 Then
    LegalCharacter = True
    Exit Function
End If


If KeyAscii < 32 Or KeyAscii = 44 Then
    LegalCharacter = False
    Exit Function
End If

If KeyAscii > 126 Then
    LegalCharacter = False
    Exit Function
End If


If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
    LegalCharacter = False
    Exit Function
End If


LegalCharacter = True

End Function
Public Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

RandomNumber = Fix(Rnd * (UpperBound - LowerBound + 1)) + LowerBound

End Function
Public Function TiempoTranscurrido(ByVal Desde As Single) As Single

TiempoTranscurrido = Timer - Desde

If TiempoTranscurrido < -5 Then
    TiempoTranscurrido = TiempoTranscurrido + 86400
ElseIf TiempoTranscurrido < 0 Then
    TiempoTranscurrido = 0
End If

End Function
Public Sub ProcesaEntradaCmd(ByVal Datos As String)

If Len(Datos) = 0 Then Exit Sub

If UCase$(left$(Datos, 3)) = "/GM" Then
    frmMSG.Show
    Exit Sub
End If

Select Case left$(Datos, 1)
    Case "\", "/"
    
    Case Else
        Datos = ";" & left$(frmMain.modo, 1) & Datos

End Select

Call SendData(Datos)

End Sub
Public Sub ResetIgnorados()
Dim i As Integer

For i = 1 To UBound(Ignorados)
    Ignorados(i) = ""
Next

End Sub
Public Function EstaIgnorado(CharIndex As Integer) As Boolean
Dim i As Integer

For i = 1 To UBound(Ignorados)
    If Len(Ignorados(i)) > 0 And Ignorados(i) = CharList(CharIndex).Nombre Then
        EstaIgnorado = True
        Exit Function
    End If
Next

End Function
Sub CheckKeys()
On Error Resume Next

Static KeyTimer As Integer

If KeyTimer > 0 Then
    KeyTimer = KeyTimer - 1
    Exit Sub
End If

If Comerciando > 0 Then Exit Sub
        
If UserMoving = 0 Then
    If Not UserEstupido Then
        If GetKeyState(vbKeyUp) < 0 Then
            Call MoveMe(NORTH)
            Exit Sub
        End If
    
        If GetKeyState(vbKeyRight) < 0 And GetKeyState(vbKeyShift) >= 0 Then
            Call MoveMe(EAST)
            Exit Sub
        End If
    
        If GetKeyState(vbKeyDown) < 0 Then
            Call MoveMe(SOUTH)
            Exit Sub
        End If

        If GetKeyState(vbKeyLeft) < 0 And GetKeyState(vbKeyShift) >= 0 Then
              Call MoveMe(WEST)
              Exit Sub
        End If
    Else
        Dim kp As Boolean
        kp = (GetKeyState(vbKeyUp) < 0) Or _
        GetKeyState(vbKeyRight) < 0 Or _
        GetKeyState(vbKeyDown) < 0 Or _
        GetKeyState(vbKeyLeft) < 0
        If kp Then Call MoveMe(Int(RandomNumber(1, 4)))
    End If
End If

End Sub
Public Function ReadField(POS As Integer, Text As String, SepASCII As Integer) As String
Dim i As Integer, LastPos As Integer, FieldNum As Integer

For i = 1 To Len(Text)
    If mid(Text, i, 1) = Chr(SepASCII) Then
        FieldNum = FieldNum + 1
        If FieldNum = POS Then
            ReadField = mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr(SepASCII), vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next

If FieldNum + 1 = POS Then ReadField = mid(Text, LastPos + 1)

End Function
Public Function PonerPuntos(Numero As Long) As String
Dim i As Integer
Dim Cifra As String

Cifra = Str(Numero)
Cifra = Right$(Cifra, Len(Cifra) - 1)
For i = 0 To 4
    If Len(Cifra) - 3 * i >= 3 Then
        If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
            PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos
        End If
    Else
        If Len(Cifra) - 3 * i > 0 Then
            PonerPuntos = left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos
        End If
        Exit For
    End If
Next

PonerPuntos = left$(PonerPuntos, Len(PonerPuntos) - 1)

End Function
Function FileExist(file As String, FileType As VbFileAttribute) As Boolean

FileExist = Len(Dir$(file, FileType)) > 0

End Function

Sub WriteClientVer()

Dim hFile As Integer
    
hFile = FreeFile()
Open App.Path & "\init\Ver.bin" For Binary Access Write As #hFile
Put #hFile, , CLng(777)
Put #hFile, , CLng(777)
Put #hFile, , CLng(777)

Put #hFile, , CInt(App.Major)
Put #hFile, , CInt(App.Minor)
Put #hFile, , CInt(App.Revision)

Close #hFile

End Sub
Function Traduccion(Original As String) As String
Dim i As Integer, Char As Integer

For i = 1 To Len(Original)
    Char = Asc(mid$(Original, i, 1)) - 232 - i ^ 2
    Do Until Char > 0
        Char = Char + 255
    Loop
    Traduccion = Traduccion & Chr$(Char)
Next
    
End Function
Sub CargarMensajes()
Dim i As Integer, NumMensajes As Integer, Leng As Byte

Open App.Path & "\Init\Mensajes.dat" For Binary As #1
Seek #1, 1

Get #1, , NumMensajes

ReDim Mensajes(1 To NumMensajes) As Mensajito

For i = 1 To NumMensajes
    Mensajes(i).Code = Space$(2)
    Get #1, , Mensajes(i).Code
    Mensajes(i).Code = Traduccion(Mensajes(i).Code)
    
    Get #1, , Leng
    Mensajes(i).mensaje = Space$(Leng)
    Get #1, , Mensajes(i).mensaje
    Mensajes(i).mensaje = Traduccion(Mensajes(i).mensaje)
    
    Get #1, , Mensajes(i).Red
    Get #1, , Mensajes(i).Green
    Get #1, , Mensajes(i).Blue
    Get #1, , Mensajes(i).Bold
    Get #1, , Mensajes(i).Italic
Next

Close #1

End Sub
Public Sub ActualizarInformacionComercio(Index As Integer)

Select Case Index
    Case 0
        frmComerciar.Label1(0).Caption = PonerPuntos(OtherInventory(frmComerciar.List1(0).ListIndex + 1).Valor)
        If OtherInventory(frmComerciar.List1(0).ListIndex + 1).Amount <> 0 Then
            frmComerciar.Label1(1).Caption = PonerPuntos(CLng(OtherInventory(frmComerciar.List1(0).ListIndex + 1).Amount))
        ElseIf OtherInventory(frmComerciar.List1(0).ListIndex + 1).Name <> "Nada" Then
            frmComerciar.Label1(1).Caption = "Ilimitado"
        Else
            frmComerciar.Label1(1).Caption = 0
        End If
        
        frmComerciar.Label1(5).Caption = OtherInventory(frmComerciar.List1(0).ListIndex + 1).Name
        frmComerciar.List1(0).ToolTipText = OtherInventory(frmComerciar.List1(0).ListIndex + 1).Name
        
        Select Case OtherInventory(frmComerciar.List1(0).ListIndex + 1).ObjType
            Case 2
                frmComerciar.Label1(3).Caption = "Max Golpe:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxHit
                frmComerciar.Label1(4).Caption = "Min Golpe:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MinHit
                frmComerciar.Label1(3).Visible = True
                frmComerciar.Label1(4).Visible = True
                frmComerciar.Label1(2).Caption = "Arma:"
                frmComerciar.Label1(2).Visible = True
            Case 3
                frmComerciar.Label1(3).Visible = True
                      frmComerciar.Label1(3).Caption = "Defensa máxima: " & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxDef
                frmComerciar.Label1(4).Caption = "Defensa mínima: " & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MinDef
                frmComerciar.Label1(4).Visible = True
                frmComerciar.Label1(2).Visible = True
                frmComerciar.Label1(2).Caption = "Casco/Escudo/Armadura"
                If OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxDef = 0 Then
                frmComerciar.Label1(3).Visible = False
                frmComerciar.Label1(4).Caption = "Esta ropa no tiene defensa."
                End If
                If OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxDef > 0 Then
                    frmComerciar.Label1(3).Visible = False
                    frmComerciar.Label1(4).Caption = "Defensa: " & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MinDef & "/" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxDef
                End If
            Case 11
                frmComerciar.Label1(3).Caption = "Max Efecto:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxModificador
                frmComerciar.Label1(4).Caption = "Min Efecto:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MinModificador
                
                frmComerciar.Label1(3).Visible = True
                frmComerciar.Label1(4).Visible = True
                frmComerciar.Label1(2).Visible = True
                frmComerciar.Label1(2).Caption = "Min Efecto:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).TipoPocion
                Select Case OtherInventory(frmComerciar.List1(0).ListIndex + 1).TipoPocion
                    Case 1
                        frmComerciar.Label1(2).Caption = "Modifica Agilidad:"
                    Case 2
                        frmComerciar.Label1(2).Caption = "Modifica Fuerza:"
                    Case 3
                        frmComerciar.Label1(2).Caption = "Repone Vida:"
                    Case 4
                        frmComerciar.Label1(2).Caption = "Repone Mana:"
                    Case 5
                        frmComerciar.Label1(2).Caption = "- Cura Envenenamiento -"
                        frmComerciar.Label1(3).Visible = False
                        frmComerciar.Label1(4).Visible = False
                End Select
            Case 24
                frmComerciar.Label1(3).Visible = False
                frmComerciar.Label1(4).Visible = False
                frmComerciar.Label1(2).Visible = True
                frmComerciar.Label1(2).Caption = "- Hechizo -"
            Case 31
                frmComerciar.Label1(3).Visible = True
                frmComerciar.Label1(4).Visible = True
                frmComerciar.Label1(2).Visible = True
                frmComerciar.Label1(2).Caption = "- Fragata -"
                frmComerciar.Label1(4).Caption = "Min/Max Golpe: " & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MinHit & "/" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxHit
                frmComerciar.Label1(3).Caption = "Defensa:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).Def
                frmComerciar.Label1(4).Visible = True
            Case Else
                frmComerciar.Label1(2).Visible = False
                frmComerciar.Label1(3).Visible = False
                frmComerciar.Label1(4).Visible = False
        End Select
        
        If OtherInventory(frmComerciar.List1(0).ListIndex + 1).PuedeUsar > 0 Then
            frmComerciar.Label1(6).Caption = "No podés usarlo ("
            Select Case OtherInventory(frmComerciar.List1(0).ListIndex + 1).PuedeUsar
                Case 1
                    frmComerciar.Label1(6).Caption = frmComerciar.Label1(6).Caption & "Genero)"
                Case 2
                    frmComerciar.Label1(6).Caption = frmComerciar.Label1(6).Caption & "Clase)"
                Case 3
                    frmComerciar.Label1(6).Caption = frmComerciar.Label1(6).Caption & "Facción)"
                Case 4
                    frmComerciar.Label1(6).Caption = frmComerciar.Label1(6).Caption & "Skill)"
                Case 5
                    frmComerciar.Label1(6).Caption = frmComerciar.Label1(6).Caption & "Raza)"
            End Select
        Else
            frmComerciar.Label1(6).Caption = ""
        End If
        
        If OtherInventory(frmComerciar.List1(0).ListIndex + 1).GrhIndex > 0 Then
            Call DrawGrhtoHdc(frmComerciar.Picture1.hdc, OtherInventory(frmComerciar.List1(0).ListIndex + 1).GrhIndex)
        Else
            frmComerciar.Picture1.Picture = LoadPicture()
        End If
        
    Case 1
        frmComerciar.Label1(0).Caption = PonerPuntos(UserInventory(frmComerciar.List1(1).ListIndex + 1).Valor)
        frmComerciar.Label1(1).Caption = PonerPuntos(UserInventory(frmComerciar.List1(1).ListIndex + 1).Amount)
        frmComerciar.Label1(5).Caption = UserInventory(frmComerciar.List1(1).ListIndex + 1).Name

        frmComerciar.List1(1).ToolTipText = UserInventory(frmComerciar.List1(1).ListIndex + 1).Name
        Select Case UserInventory(frmComerciar.List1(1).ListIndex + 1).ObjType
            Case 2
                frmComerciar.Label1(2).Caption = "Arma:"
                frmComerciar.Label1(3).Caption = "Max Golpe:" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxHit
                frmComerciar.Label1(4).Caption = "Min Golpe:" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MinHit
                frmComerciar.Label1(3).Visible = True
                frmComerciar.Label1(2).Visible = True
                frmComerciar.Label1(4).Visible = True
            Case 3
                frmComerciar.Label1(3).Visible = True
                frmComerciar.Label1(3).Caption = "Defensa máxima: " & UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxDef
                frmComerciar.Label1(4).Caption = "Defensa mínima: " & UserInventory(frmComerciar.List1(1).ListIndex + 1).MinDef
                frmComerciar.Label1(4).Visible = True
                frmComerciar.Label1(2).Visible = True
                frmComerciar.Label1(2).Caption = "Casco/Escudo/Armadura"
                If UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxDef = 0 Then
                frmComerciar.Label1(3).Visible = False
                frmComerciar.Label1(4).Caption = "Esta ropa no tiene defensa."
                End If
                If UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxDef > 0 Then
                    frmComerciar.Label1(3).Visible = False
                    frmComerciar.Label1(4).Caption = "Defensa " & UserInventory(frmComerciar.List1(1).ListIndex + 1).MinDef & "/" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxDef
                End If
            Case 11
                frmComerciar.Label1(3).Caption = "Max Efecto:" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxModificador
                frmComerciar.Label1(4).Caption = "Min Efecto:" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MinModificador
                
                frmComerciar.Label1(3).Visible = True
                frmComerciar.Label1(4).Visible = True
                frmComerciar.Label1(2).Visible = True
                
                Select Case UserInventory(frmComerciar.List1(1).ListIndex + 1).TipoPocion
                    Case 1
                        frmComerciar.Label1(2).Caption = "Aumenta Agilidad"
                    Case 2
                        frmComerciar.Label1(2).Caption = "Aumenta Fuerza"
                    Case 3
                        frmComerciar.Label1(2).Caption = "Repone Vida"
                    Case 4
                        frmComerciar.Label1(2).Caption = "Repone Mana"
                    Case 5
                        frmComerciar.Label1(2).Caption = "- Cura Envenenamiento -"
                        frmComerciar.Label1(3).Visible = False
                        frmComerciar.Label1(4).Visible = False
                End Select
            Case 24
                frmComerciar.Label1(3).Visible = False
                frmComerciar.Label1(4).Visible = False
                frmComerciar.Label1(2).Caption = "- Hechizo -"
                frmComerciar.Label1(2).Visible = True
            Case 31
                frmComerciar.Label1(3).Visible = True
                frmComerciar.Label1(4).Visible = True
                frmComerciar.Label1(2).Caption = "- Fragata -"
                frmComerciar.Label1(4).Caption = "Min/Max Golpe: " & UserInventory(frmComerciar.List1(1).ListIndex + 1).MinHit & "/" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxHit
                frmComerciar.Label1(3).Caption = "Defensa:" & UserInventory(frmComerciar.List1(1).ListIndex + 1).Def
                frmComerciar.Label1(4).Visible = True
            frmComerciar.Label1(2).Visible = True
            Case Else
                frmComerciar.Label1(2).Visible = False
                frmComerciar.Label1(3).Visible = False
                frmComerciar.Label1(4).Visible = False
        End Select
        
        If UserInventory(frmComerciar.List1(1).ListIndex + 1).GrhIndex > 0 Then
            Call DrawGrhtoHdc(frmComerciar.Picture1.hdc, UserInventory(frmComerciar.List1(1).ListIndex + 1).GrhIndex)
        Else
            frmComerciar.Picture1.Picture = LoadPicture()
        End If
        
End Select

frmComerciar.Picture1.Refresh

End Sub
Sub TelepPorMapa(X As Long, Y As Long)
Dim Columna As Long, Fila As Long

Columna = Fix((X - 25) / 18)
Fila = Fix((Y - 18) / 18)

Call SendData("#$" & Columna & "," & Fila)

End Sub

Sub Main()
'On Error Resume Next


FrmIntro.Hide

AddtoRichTextBox frmCargando.Status, "Cargando FenixAO...", 255, 150, 50, 1, , False

Call WriteClientVer

CartelOcultarse = Val(GetVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Ocultarse"))
CartelMenosCansado = Val(GetVar(App.Path & "/Init/Opciones.opc", "CARTELES", "MenosCansado"))
CartelVestirse = Val(GetVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Vestirse"))
CartelNoHayNada = Val(GetVar(App.Path & "/Init/Opciones.opc", "CARTELES", "NoHayNada"))
CartelRecuMana = Val(GetVar(App.Path & "/Init/Opciones.opc", "CARTELES", "RecuMana"))
CartelSanado = Val(GetVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Sanado"))
NoRes = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "ModoVentana"))
ResOriginal = NoRes
Musica = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "Musica"))
FX = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "FX"))

If ResOriginal <> True Then Call SetResolution

If App.PrevInstance Then
    Call MsgBox("¡FénixAO ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
    End
End If

ChDrive App.Path
ChDir App.Path

frmCargando.Show
frmCargando.Refresh

UserParalizado = False

AddtoRichTextBox frmCargando.Status, "Iniciando constantes...", 255, 150, 50, 0, , True

RG(1, 1) = 255
RG(1, 2) = 128
RG(1, 3) = 64

RG(2, 1) = 0
RG(2, 2) = 128
RG(2, 3) = 255

RG(3, 1) = 255
RG(3, 2) = 0
RG(3, 3) = 0

RG(4, 1) = 0
RG(4, 2) = 240
RG(4, 3) = 0

RG(5, 1) = 190
RG(5, 2) = 190
RG(5, 3) = 190

ReDim Ciudades(1 To NUMCIUDADES) As String
Ciudades(1) = "Ullathorpe"
Ciudades(2) = "Nix"
Ciudades(3) = "Banderbill"

ReDim CityDesc(1 To NUMCIUDADES) As String
CityDesc(1) = "Ullathorpe está establecida en el medio de los grandes bosques de Argentum, es principalmente un pueblo de campesinos y leñadores. Su ubicación hace de Ullathorpe un punto de paso obligado para todos los aventureros ya que se encuentra cerca de los lugares más legendarios de este mundo."
CityDesc(2) = "Nix es una gran ciudad. Edificada sobre la costa oeste del principal continente de Argentum."
CityDesc(3) = "Banderbill se encuentra al norte de Ullathorpe y Nix, es una de las ciudades más importantes de todo el imperio."

ReDim ListaRazas(1 To NUMRAZAS) As String
ListaRazas(1) = "Humano"
ListaRazas(2) = "Elfo"
ListaRazas(3) = "Elfo Oscuro"
ListaRazas(4) = "Gnomo"
ListaRazas(5) = "Enano"

ReDim ListaClases(1 To NUMCLASES) As String
ListaClases(1) = "Mago"
ListaClases(2) = "Clerigo"
ListaClases(3) = "Guerrero"
ListaClases(4) = "Asesino"
ListaClases(5) = "Ladron"
ListaClases(6) = "Bardo"
ListaClases(7) = "Druida"
ListaClases(8) = "Bandido"
ListaClases(9) = "Paladin"
ListaClases(10) = "Arquero"
ListaClases(11) = "Pescador"
ListaClases(12) = "Herrero"
ListaClases(13) = "Leñador"
ListaClases(14) = "Minero"
ListaClases(15) = "Carpintero"
ListaClases(16) = "Pirata"

ReDim SkillsNames(1 To NUMSKILLS) As String
SkillsNames(1) = "Magia"
SkillsNames(2) = "Robar"
SkillsNames(3) = "Tacticas de combate"
SkillsNames(4) = "Combate con armas"
SkillsNames(5) = "Meditar"
SkillsNames(6) = "Apuñalar"
SkillsNames(7) = "Ocultarse"
SkillsNames(8) = "Supervivencia"
SkillsNames(9) = "Talar árboles"
SkillsNames(10) = "Defensa con escudos"
SkillsNames(11) = "Pesca"
SkillsNames(12) = "Mineria"
SkillsNames(13) = "Carpinteria"
SkillsNames(14) = "Herreria"
SkillsNames(15) = "Liderazgo"
SkillsNames(16) = "Domar animales"
SkillsNames(17) = "Armas de proyectiles"
SkillsNames(18) = "Wresterling"
SkillsNames(19) = "Navegacion"
SkillsNames(20) = "Sastrería"
SkillsNames(21) = "Comercio"
SkillsNames(22) = "Resistencia Mágica"

ReDim UserSkills(1 To NUMSKILLS) As Integer
ReDim UserAtributos(1 To NUMATRIBUTOS) As Integer
ReDim AtributosNames(1 To NUMATRIBUTOS) As String
AtributosNames(1) = "Fuerza"
AtributosNames(2) = "Agilidad"
AtributosNames(3) = "Inteligencia"
AtributosNames(4) = "Carisma"
AtributosNames(5) = "Constitucion"

AddtoRichTextBox frmCargando.Status, "Hecho", 255, 150, 50, 1, , False

AddtoRichTextBox frmCargando.Status, "Cargando motor de sonido....", 255, 150, 50, , , True

'Inicializamos el sonido
    Call Audio.Initialize(frmMain.hWnd, App.Path & "\Wav\", App.Path & "\Midi\")

AddtoRichTextBox frmCargando.Status, "Hecho", 255, 150, 50, 1, , False

ENDC = Chr(1)

UserMap = 1

AddtoRichTextBox frmCargando.Status, "Cargando motor grafico....", 255, 150, 50, , , True

Call InitTileEngine(frmMain.Renderer.hWnd, 32, 32, 12, 16)


AddtoRichTextBox frmCargando.Status, "Hecho", 255, 150, 50, 1, , False

Call AddtoRichTextBox(frmCargando.Status, "Creando animaciones extras.", 255, 150, 50, 1, , True)

Unload frmCargando

Call Audio.PlayMIDI(MIdi_Inicio & ".mid")

frmPres.Picture = LoadPicture(App.Path & "\Graficos\fenix.jpg")
frmPres.WindowState = vbMaximized
frmPres.Show

Do While Not finpres
    DoEvents
Loop

Unload frmPres


frmConnect.Visible = True

PrimeraVez = True
prgRun = True
Pausa = False

' Empieza el bucle
Call ShowNextFrame

EngineRun = False
frmCargando.Show
AddtoRichTextBox frmCargando.Status, "Liberando recursos...", 0, 0, 0, 0, 0, 1

If ResOriginal <> True Then Call ResetResolution

Call UnloadAllForms
Call DeInitTileEngine

End

'ManejadorErrores:
'    End
    
End Sub



Sub WriteVar(file As String, Main As String, Var As String, value As String)


writeprivateprofilestring Main, Var, value, file

End Sub

Function GetVar(file As String, Main As String, Var As String) As String
Dim l As Integer
Dim Char As String
Dim sSpaces As String
Dim szReturn As String

szReturn = ""

sSpaces = Space(5000)


getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), file

GetVar = RTrim(sSpaces)
GetVar = left$(GetVar, Len(GetVar) - 1)

End Function
Public Function CheckMailString(ByRef sString As String) As Boolean
On Error GoTo errHnd:
Dim lPos As Long, lX As Long

lPos = InStr(sString, "@")
If (lPos <> 0) Then
    If Not InStr(lPos, sString, ".", vbBinaryCompare) > (lPos + 1) Then Exit Function

    For lX = 0 To Len(sString) - 1
        If Not lX = (lPos - 1) And Not CMSValidateChar_(Asc(mid$(sString, (lX + 1), 1))) Then Exit Function
    Next lX

    CheckMailString = True
End If
    
errHnd:

End Function
Private Function CMSValidateChar_(ByRef iAsc As Integer) As Boolean

CMSValidateChar_ = iAsc = 46 Or (iAsc >= 48 And iAsc <= 57) Or _
                    (iAsc >= 65 And iAsc <= 90) Or _
                    (iAsc >= 97 And iAsc <= 122) Or _
                    (iAsc = 95) Or (iAsc = 45)
                    
End Function

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Integer, ByRef tY As Integer)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.X + viewPortX \ 32 - frmMain.Renderer.ScaleWidth \ 64
    tY = UserPos.Y + viewPortY \ 32 - frmMain.Renderer.ScaleHeight \ 64
    Debug.Print tX; tY
End Sub

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        count = count + 1
    Loop While curPos <> 0
    
    FieldCount = count
End Function
