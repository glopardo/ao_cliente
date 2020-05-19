Attribute VB_Name = "Mod_General"
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


Public CartelOcultarse As Byte
Public CartelMenosCansado As Byte
Public CartelVestirse As Byte
Public CartelNoHayNada As Byte
Public CartelRecuMana As Byte
Public CartelSanado As Byte
Public atacar As Integer
Public IsClan As Byte
Public NoRes As Boolean
Public Desplazar As Boolean
Public vigilar As Boolean


Public RG(1 To 5, 1 To 3) As Byte

Public bO As Integer
Public bK As Long
Public bRK As Long
Public iplst As String
Public banners As String

Public bInvMod     As Boolean

Public bFogata As Boolean

Public bLluvia() As Byte

Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal wIndx As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Private lFrameLimiter As Long

Public lFrameModLimiter As Long
Public lFrameTimer As Long
Public sHKeys() As String

Public bFPS As Boolean
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal _
dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
ByVal dwProcessId As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject _
As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" _
   (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long

Const PROCESS_TERMINATE = &H1
Const PROCESS_QUERY_INFORMATION = &H400
Const STILL_ACTIVE = &H103

Type Recompensa
    Name As String
    Descripcion As String
End Type

Const GWL_STYLE = (-16)
Const Win_VISIBLE = &H10000000
Const Win_BORDER = &H800000
Const SC_CLOSE = &HF060&
Const WM_SYSCOMMAND = &H112

Dim ObjetoWMI As Object
Dim ProcesoACerrar As Object
Dim Procesos As Object
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
DirGraficos = App.Path & "\" & Config_Inicio.DirGraficos & "\"
End Function

Public Function DirSound() As String
DirSound = App.Path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
DirMidi = App.Path & "\" & Config_Inicio.DirMusica & "\"
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
Sub PlayWaveAPI(File As String)

On Error Resume Next
Dim rc As Integer

rc = sndPlaySound(File, SND_ASYNC)

End Sub
Sub CargarAnimArmas()

On Error Resume Next

Dim loopc As Integer
Dim arch As String
arch = App.Path & "\init\" & "armas.dat"
DoEvents

NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))

ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData

For loopc = 1 To NumWeaponAnims
    InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
    InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
    InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
    InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
Next loopc

End Sub
Sub CargarAnimEscudos()
On Error Resume Next

Dim loopc As Integer
Dim arch As String
arch = App.Path & "\init\" & "escudos.dat"
DoEvents

NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))

ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData

For loopc = 1 To NumEscudosAnims
    InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
    InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
    InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
    InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
Next loopc

End Sub

Sub Addtostatus(RichTextBox As RichTextBox, Text As String, Red As Byte, Green As Byte, Blue As Byte, Bold As Byte, Italic As Byte)







frmCargando.Status.SelStart = Len(RichTextBox.Text)
frmCargando.Status.SelLength = 0
frmCargando.Status.SelColor = RGB(Red, Green, Blue)

If Bold Then
    frmCargando.Status.SelBold = True
Else
    frmCargando.Status.SelBold = False
End If

If Italic Then
    frmCargando.Status.SelItalic = True
Else
    frmCargando.Status.SelItalic = False
End If

frmCargando.Status.SelText = Chr(13) & Chr(10) & Text

End Sub
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
Sub RefreshAllChars()
Dim loopc As Integer

For loopc = 1 To LastChar
    If CharList(loopc).Active = 1 Then
        MapData(CharList(loopc).POS.X, CharList(loopc).POS.Y).CharIndex = loopc
    End If
Next loopc

End Sub
Sub SaveGameini()

Config_Inicio.Name = "BetaTester"
Config_Inicio.Password = "DammLamers"
Config_Inicio.Puerto = UserPort

Call EscribirGameIni(Config_Inicio)

End Sub
Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    
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
    CharAscii = Asc(Mid$(UserPassword, loopc, 1))
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

    CharAscii = Asc(Mid$(UserName, loopc, 1))
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

Sub SetConnected()





Connected = True

Call SaveGameini


Unload frmConnect


frmMain.Label8.Caption = UserName

frmMain.Visible = True



End Sub
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
Sub MoveMe(Direction As Byte)

If CONGELADO Then Exit Sub

If Cartel Then Cartel = False

If ProxLegalPos(Direction) And Not UserMeditar And Not UserParalizado Then
    If TiempoTranscurrido(LastPaso) >= IntervaloPaso Then
        Call SendData("M" & Direction)
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

frmMain.mapa.Caption = NombreDelMapaActual & " [" & UserMap & " - " & UserPos.X & " - " & UserPos.Y & "]"

End Sub
Function ProxLegalPos(Direction As Byte) As Boolean

Select Case Direction
    Case NORTH
        ProxLegalPos = LegalPos(UserPos.X, UserPos.Y - 1)
    Case SOUTH
        ProxLegalPos = LegalPos(UserPos.X, UserPos.Y + 1)
    Case WEST
        ProxLegalPos = LegalPos(UserPos.X - 1, UserPos.Y)
    Case EAST
        ProxLegalPos = LegalPos(UserPos.X + 1, UserPos.Y)
End Select

End Function
Sub EliminarChars(Direction As Byte)
Dim X(2) As Integer
Dim Y(2) As Integer

Select Case Direction
    Case NORTH, SOUTH
        X(1) = UserPos.X - MinXBorder - 2
        X(2) = UserPos.X + MinXBorder + 2
    Case EAST, WEST
        Y(1) = UserPos.Y - MinYBorder - 2
        Y(2) = UserPos.Y + MinYBorder + 2
End Select

Select Case Direction
    Case NORTH
        Y(1) = UserPos.Y - MinYBorder - 3
        If Y(1) < 1 Then Y(1) = 1
        Y(2) = Y(1)
    Case EAST
        X(1) = UserPos.X + MinXBorder + 3
        If X(1) > 99 Then X(1) = 99
        X(2) = X(1)
    Case SOUTH
        Y(1) = UserPos.Y + MinYBorder + 3
        If Y(1) > 99 Then Y(1) = 99
        Y(2) = Y(1)
    Case WEST
        X(1) = UserPos.X - MinXBorder - 3
        If X(1) < 1 Then X(1) = 1
        X(2) = X(1)
End Select

For Y(0) = Y(1) To Y(2)
    For X(0) = X(1) To X(2)
        If X(0) > 6 And X(0) < 95 And Y(0) > 6 And Y(0) < 95 Then
            If MapData(X(0), Y(0)).CharIndex > 0 Then
                CharList(MapData(X(0), Y(0)).CharIndex).POS.X = 0
                CharList(MapData(X(0), Y(0)).CharIndex).POS.Y = 0
                MapData(X(0), Y(0)).CharIndex = 0
            End If
        End If
    Next
Next

End Sub
Public Sub ProcesaEntradaCmd(ByVal Datos As String)

If Len(Datos) = 0 Then Exit Sub

If UCase$(Left$(Datos, 3)) = "/GM" Then
    frmMSG.Show
    Exit Sub
End If

Select Case Left$(Datos, 1)
    Case "\", "/"
    
    Case Else
        Datos = ";" & Left$(frmMain.modo, 1) & Datos

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
Sub MoveScreen(Heading As Byte)
Dim X As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer
Dim bx As Integer
Dim by As Integer

Select Case Heading

    Case NORTH
        Y = -1

    Case EAST
        X = 1

    Case SOUTH
        Y = 1
    
    Case WEST
        X = -1
        
End Select

tX = UserPos.X + X
tY = UserPos.Y + Y



If Not (tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder) Then
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1

    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
Exit Sub
Stop
    
        
        Select Case FramesPerSecCounter
            Case Is >= 17
                lFrameModLimiter = 60
            Case 16
                lFrameModLimiter = 120
            Case 15
                lFrameModLimiter = 240
            Case 14
                lFrameModLimiter = 480
            Case 15
                lFrameModLimiter = 960
            Case 14
                lFrameModLimiter = 1920
            Case 13
                lFrameModLimiter = 3840
            Case 1
                lFrameModLimiter = 60 * 256
            Case 0
            
        End Select
    

    Call DoFogataFx
End If

End Sub
Function NextOpenChar()
Dim loopc As Integer

loopc = 1

Do While CharList(loopc).Active
    loopc = loopc + 1
Loop

NextOpenChar = loopc

End Function
Public Function DirMapas() As String

DirMapas = App.Path & "\maps\"

End Function
Sub SwitchMap(Map As Integer)
Dim loopc As Integer
Dim Y As Integer
Dim X As Integer
Dim tempint As Integer

Open DirMapas & "Mapa" & Map & ".mcl" For Binary As #1
Seek #1, 1
        

Get #1, , MapInfo.MapVersion
Get #1, , MiCabecera
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
        

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        
        Get #1, , MapData(X, Y).Blocked
        For loopc = 1 To 4
            Get #1, , MapData(X, Y).Graphic(loopc).GrhIndex
            If loopc = 3 And MapData(X, Y).Graphic(loopc).GrhIndex <> 0 And MapData(X, Y).Blocked = 1 Then
                MapData(X, Y).ObjGrh = MapData(X, Y).Graphic(loopc)
                MapData(X, Y).Graphic(loopc).GrhIndex = 0
            End If
            
            If MapData(X, Y).Graphic(loopc).GrhIndex > 0 Then
                If MapData(X, Y).Graphic(loopc).GrhIndex = 7000 Then MapData(X, Y).Graphic(loopc).GrhIndex = 700
                InitGrh MapData(X, Y).Graphic(loopc), MapData(X, Y).Graphic(loopc).GrhIndex
            End If
            
        Next loopc
        
        
        Get #1, , MapData(X, Y).Trigger
        
        Get #1, , tempint
        
        
        If MapData(X, Y).CharIndex > 0 Then
            Call EraseChar(MapData(X, Y).CharIndex)
        End If
        
        
        MapData(X, Y).ObjGrh.GrhIndex = 0

    Next X
Next Y

Close #1

MapInfo.Name = ""
MapInfo.Music = ""
CurMap = Map

End Sub
Sub EliminarDatosMapa()
Dim X As Integer
Dim Y As Integer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(X, Y).CharIndex > 0 Then Call EraseChar(MapData(X, Y).CharIndex)
        MapData(X, Y).ObjGrh.GrhIndex = 0
    Next X
Next Y

End Sub
Sub SwitchMapBaley(Map As Integer)
On Error Resume Next
Dim loopc As Integer
Dim Y As Integer
Dim X As Integer
Dim tempint As Integer
Dim InfoTile As Byte
Dim i As Integer

Open DirMapas & "Mapa" & Map & ".mcl" For Binary As #1
Seek #1, 1
        

Get #1, , MapInfo.MapVersion
Get #1, , MiCabecera

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        Get #1, , InfoTile
        
        MapData(X, Y).Blocked = (InfoTile And 1)
        
        Get #1, , MapData(X, Y).Graphic(1).GrhIndex
        Call InitGrh(MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex)
        
        For i = 2 To 4
            If InfoTile And (2 ^ (i - 1)) Then
                Get #1, , MapData(X, Y).Graphic(i).GrhIndex
                If i = 3 And MapData(X, Y).Graphic(3).GrhIndex <> 0 And MapData(X, Y).Blocked = 1 Then
                    MapData(X, Y).ObjGrh = MapData(X, Y).Graphic(3)
                    MapData(X, Y).Graphic(3).GrhIndex = 0
                End If
                    
                Call InitGrh(MapData(X, Y).Graphic(i), MapData(X, Y).Graphic(i).GrhIndex)

            Else
                MapData(X, Y).Graphic(i).GrhIndex = 0
            End If
        Next
        
        If InfoTile And 16 Then Get #1, , MapData(X, Y).Trigger
        
        If MapData(X, Y).CharIndex > 0 Then Call EraseChar(MapData(X, Y).CharIndex)

        MapData(X, Y).ObjGrh.GrhIndex = 0

    Next X
Next Y

Close #1

MapInfo.Name = ""

MapInfo.Music = ""

CurMap = Map

End Sub
Sub SwitchMapNew(Map As Integer)
On Error Resume Next
Dim loopc As Integer
Dim Y As Integer
Dim X As Integer
Dim tempint As Integer
Dim InfoTile As Byte
Dim i As Integer

Open DirMapas & "Mapa" & Map & ".mcl" For Binary As #1
Seek #1, 1
        

Get #1, , MapInfo.MapVersion

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        Get #1, , InfoTile
        
        MapData(X, Y).Blocked = (InfoTile And 1)
        
        Get #1, , MapData(X, Y).Graphic(1).GrhIndex
        
        For i = 2 To 4
            If InfoTile And (2 ^ (i - 1)) Then
                Get #1, , MapData(X, Y).Graphic(i).GrhIndex
                Call InitGrh(MapData(X, Y).Graphic(i), MapData(X, Y).Graphic(i).GrhIndex)
            Else: MapData(X, Y).Graphic(i).GrhIndex = 0
            End If
        Next
        
        MapData(X, Y).Trigger = 0
        
        For i = 4 To 6
            If (InfoTile And 2 ^ i) Then MapData(X, Y).Trigger = MapData(X, Y).Trigger Or 2 ^ (i - 4)
        Next
        
        Call InitGrh(MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex)
    
        If MapData(X, Y).CharIndex > 0 Then Call EraseChar(MapData(X, Y).CharIndex)
        MapData(X, Y).ObjGrh.GrhIndex = 0
    Next X
Next Y

Close #1

MapInfo.Name = ""

MapInfo.Music = ""

CurMap = Map

End Sub
Public Function ReadField(POS As Integer, Text As String, SepASCII As Integer) As String
Dim i As Integer, LastPos As Integer, FieldNum As Integer

For i = 1 To Len(Text)
    If Mid(Text, i, 1) = Chr(SepASCII) Then
        FieldNum = FieldNum + 1
        If FieldNum = POS Then
            ReadField = Mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr(SepASCII), vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next

If FieldNum + 1 = POS Then ReadField = Mid(Text, LastPos + 1)

End Function
Public Function NumeroApuesta(Numero As Integer) As String
Dim MiNum As Byte

Select Case Numero
    Case Is <= 36
        NumeroApuesta = "l " & Numero & "."
    Case 37
        NumeroApuesta = " los primeros 12."
    Case 38
        NumeroApuesta = " los segundos 12."
    Case 39
        NumeroApuesta = " los últimos 12."
    Case 40
        NumeroApuesta = " los primeros 18."
    Case 41
        NumeroApuesta = " los pares."
    Case 42
        NumeroApuesta = " los rojos."
    Case 43
        NumeroApuesta = " los negros."
    Case 44
        NumeroApuesta = " los impares."
    Case 45
        NumeroApuesta = " los últimos 18."
    Case Is <= 69
        MiNum = 3 * Fix((Numero - 46) / 2) + 2
        If Numero Mod 2 = 0 Then
            NumeroApuesta = "l semipleno " & MiNum - 1 & "-" & MiNum & "."
        Else
            NumeroApuesta = "l semipleno " & MiNum & "-" & MiNum + 1 & "."
        End If
    Case Is <= 102
        NumeroApuesta = "l semipleno " & Numero - 69 & "-" & Numero - 66 & "."
    Case Is <= 124
        MiNum = (3 * Fix((Numero - 101) / 2) - 1)
        If Numero Mod 2 = 1 Then MiNum = MiNum - 1
        NumeroApuesta = "l cuadro " & MiNum & "-" & MiNum + 1 & "-" & MiNum + 3 & "-" & MiNum + 4 & "."
    Case Is <= 136
        MiNum = 1 + 3 * (Numero - 125)
        NumeroApuesta = " la fila del " & MiNum & " al " & MiNum + 2 & "."
    Case Is <= 147
        MiNum = 1 + 3 * (Numero - 137)
        NumeroApuesta = " la calle del " & MiNum & " al " & MiNum + 5 & "."
    Case 148
        NumeroApuesta = " la primer columna."
    Case 149
        NumeroApuesta = " la segunda columna."
    Case 150
        NumeroApuesta = " la tercer columna."
End Select
        
End Function
Public Function PonerPuntos(Numero As Long) As String
Dim i As Integer
Dim Cifra As String

Cifra = Str(Numero)
Cifra = Right$(Cifra, Len(Cifra) - 1)
For i = 0 To 4
    If Len(Cifra) - 3 * i >= 3 Then
        If Mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
            PonerPuntos = Mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos
        End If
    Else
        If Len(Cifra) - 3 * i > 0 Then
            PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos
        End If
        Exit For
    End If
Next

PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)

End Function
Function FileExist(File As String, FileType As VbFileAttribute) As Boolean

FileExist = Len(Dir$(File, FileType)) > 0

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

Sub ReNombrarAutoUpdate()

If FileExist(App.Path & "\NuevoUpdater.exe", vbNormal) Then
    If FileExist(App.Path & "\AutoUpdateClient.exe", vbNormal) Then Call Kill(App.Path & "\AutoUpdateClient.exe")
    Name App.Path & "\NuevoUpdater.exe" As App.Path & "\AutoUpdateClient.exe"
End If

End Sub
Public Function IsIp(ByVal Ip As String) As Boolean

Dim i As Integer
For i = 1 To UBound(ServersLst)
    If ServersLst(i).Ip = Ip Then
        IsIp = True
        Exit Function
    End If
Next i

End Function

Public Sub InitServersList(ByVal Lst As String)

Dim NumServers As Integer
Dim i As Integer, Cont As Integer
i = 1

Do While (ReadField(i, RawServersList, Asc(";")) <> "")
    i = i + 1
    Cont = Cont + 1
Loop

ReDim ServersLst(1 To Cont) As tServerInfo

For i = 1 To Cont
    Dim cur$
    cur$ = ReadField(i, RawServersList, Asc(";"))
    ServersLst(i).Ip = ReadField(1, cur$, Asc(":"))
    ServersLst(i).Puerto = ReadField(2, cur$, Asc(":"))
    ServersLst(i).desc = ReadField(4, cur$, Asc(":"))
    ServersLst(i).PassRecPort = ReadField(3, cur$, Asc(":"))
Next i

CurServer = 1

End Sub
Sub CargarMensajesV()
Dim i As Integer
Dim File As String
Dim Formato As String
Dim NumMensajes As Integer

File = App.Path & "\Init\MensajesV.dat"

NumMensajes = Val(GetVar(File, "INIT", "NumMensajes"))

ReDim Mensajes(1 To NumMensajes) As Mensajito

For i = 1 To NumMensajes
    Mensajes(i).Code = GetVar(File, "Mensaje" & i, "C")
    Mensajes(i).mensaje = GetVar(File, "Mensaje" & i, "M")
    Formato = GetVar(File, "Mensaje" & i, "F")
    Mensajes(i).Red = Val(ReadField(1, Formato, Asc("-")))
    Mensajes(i).Green = Val(ReadField(2, Formato, Asc("-")))
    Mensajes(i).Blue = Val(ReadField(3, Formato, Asc("-")))
    Mensajes(i).Bold = Val(ReadField(4, Formato, Asc("-")))
    Mensajes(i).Italic = Val(ReadField(5, Formato, Asc("-")))
Next

Call SaveMensajes

End Sub
Function Transcripcion(Original As String) As String
Dim i As Integer, Char As Integer

For i = 1 To Len(Original)
    Char = Asc(Mid$(Original, i, 1)) + 232 + i ^ 2
    Do Until Char < 255
        Char = Char - 255
    Loop
    Transcripcion = Transcripcion & Chr$(Char)
Next
    
End Function
Function Traduccion(Original As String) As String
Dim i As Integer, Char As Integer

For i = 1 To Len(Original)
    Char = Asc(Mid$(Original, i, 1)) - 232 - i ^ 2
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
Sub SaveMensajes()
Dim i As Integer, File As String

File = App.Path & "\Init\Mensajes.dat"

Open File For Binary As #1
Seek #1, 1

Put #1, , CInt(UBound(Mensajes))
For i = 1 To UBound(Mensajes)
    Put #1, , Transcripcion(Mensajes(i).Code)
    Put #1, , CByte(Len(Mensajes(i).mensaje))
    Put #1, , Transcripcion(Mensajes(i).mensaje)
    Put #1, , Mensajes(i).Red
    Put #1, , Mensajes(i).Green
    Put #1, , Mensajes(i).Blue
    Put #1, , Mensajes(i).Bold
    Put #1, , Mensajes(i).Italic
Next

Close #1

End Sub
Public Sub ActualizarInformacionComercio(Index As Integer)

Dim SR As RECT, DR As RECT
SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.Bottom = 32

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
            Call DrawGrhtoHdc(frmComerciar.Picture1.hwnd, frmComerciar.Picture1.Hdc, OtherInventory(frmComerciar.List1(0).ListIndex + 1).GrhIndex, SR, DR)
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
            Call DrawGrhtoHdc(frmComerciar.Picture1.hwnd, frmComerciar.Picture1.Hdc, UserInventory(frmComerciar.List1(1).ListIndex + 1).GrhIndex, SR, DR)
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
On Error Resume Next


FrmIntro.Hide

AddtoRichTextBox frmCargando.Status, "Cargando...", 255, 150, 50, 1, , False

Call WriteClientVer

CartelOcultarse = Val(GetVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Ocultarse"))
CartelMenosCansado = Val(GetVar(App.Path & "/Init/Opciones.opc", "CARTELES", "MenosCansado"))
CartelVestirse = Val(GetVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Vestirse"))
CartelNoHayNada = Val(GetVar(App.Path & "/Init/Opciones.opc", "CARTELES", "NoHayNada"))
CartelRecuMana = Val(GetVar(App.Path & "/Init/Opciones.opc", "CARTELES", "RecuMana"))
CartelSanado = Val(GetVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Sanado"))
NoRes = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "ModoVentana"))

If App.PrevInstance Then
    Call MsgBox("¡Argentum Online ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
    End
End If

Dim f As Boolean
Dim ulttick As Long, esttick As Long
Dim timers(1 To 5) As Long
ChDrive App.Path
ChDir App.Path


Dim fMD5HushYo As String * 32
HushYo = GenHash(App.Path & "\" & App.exename & ".exe")


If FileExist(App.Path & "\init\Inicio.con", vbNormal) Then
    Config_Inicio = LeerGameIni()
End If

If FileExist(App.Path & "\init\ao.dat", vbNormal) Then
    Open App.Path & "\init\ao.dat" For Binary As #53
        Get #53, , RenderMod
    Close #53

    Musica = IIf(RenderMod.bNoMusic = 1, 1, 0)
    FX = IIf(RenderMod.bNoSound = 1, 1, 0)
    
    
    Select Case RenderMod.iImageSize
        Case 4
            RenderMod.iImageSize = 0
        Case 3
            RenderMod.iImageSize = 1
        Case 2
            RenderMod.iImageSize = 2
        Case 1
            RenderMod.iImageSize = 3
        Case 0
            RenderMod.iImageSize = 4
    End Select
End If


tipf = Config_Inicio.tip

frmCargando.Show
frmCargando.Refresh

UserParalizado = False

AddtoRichTextBox frmCargando.Status, "Buscando servidores....", 255, 150, 50, , , True

AddtoRichTextBox frmCargando.Status, "Encontrado", 255, 150, 50, 1, , False
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

IniciarObjetosDirectX

AddtoRichTextBox frmCargando.Status, "Cargando Sonidos....", 255, 150, 50, , , True
AddtoRichTextBox frmCargando.Status, "Hecho", 255, 150, 50, 1, , False

Dim loopc As Integer

LastTime = GetTickCount

ENDL = Chr(13) & Chr(10)
ENDC = Chr(1)


Call InitTileEngine(frmMain.hwnd, frmMain.MainViewShp.Top, frmMain.MainViewShp.Left, 32, 32, 13, 17, 9)

Call AddtoRichTextBox(frmCargando.Status, "Creando animaciones extras.", 255, 150, 50, 1, , True)

UserMap = 1

Call CargarAnimsExtra
Call CargarArrayLluvia
Call CargarAnimArmas
Call CargarAnimEscudos
Call CargarMensajes
Call EstablecerRecompensas

Unload frmCargando

LoopMidi = True

If Musica = 0 Then
    Call CargarMIDI(DirMidi & MIdi_Inicio & ".mid")
    Play_Midi
End If

frmPres.Picture = LoadPicture(App.Path & "\Graficos\fenix.jpg")
frmPres.WindowState = vbMaximized
frmPres.Show

Do While Not finpres
    DoEvents
Loop

Unload frmPres


frmConnect.Visible = True

MainViewRect.Left = (frmMain.Left / Screen.TwipsPerPixelX) + MainViewLeft + 32 * RenderMod.iImageSize
MainViewRect.Top = (frmMain.Top / Screen.TwipsPerPixelY) + MainViewTop + 32 * RenderMod.iImageSize
MainViewRect.Right = (MainViewRect.Left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
MainViewRect.Bottom = (MainViewRect.Top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)

MainDestRect.Left = ((TilePixelWidth * TileBufferSize) - TilePixelWidth) + 32 * RenderMod.iImageSize
MainDestRect.Top = ((TilePixelHeight * TileBufferSize) - TilePixelHeight) + 32 * RenderMod.iImageSize
MainDestRect.Right = (MainDestRect.Left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
MainDestRect.Bottom = (MainDestRect.Top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)

Dim OffsetCounterX As Integer
Dim OffsetCounterY As Integer

PrimeraVez = True
prgRun = True
Pausa = False
lFrameLimiter = DirectX.TickCount


Do While prgRun
        
        
        
        
        
        
        
        
    
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
    If RequestPosTimer > 0 Then
        RequestPosTimer = RequestPosTimer - 1
        If RequestPosTimer = 0 Then
            
            Call SendData("RPU")
        End If
    End If

    Call RefreshAllChars

    
    
    
    If EngineRun Then
        
        
        
        If frmMain.WindowState <> 1 Then
        
            
            
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = (OffsetCounterX - (IIf(UserMontando, (32 / 3), 8) * Sgn(AddtoUserPos.X)))
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = 0
                End If
            ElseIf AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - (IIf(UserMontando, (32 / 3), 8) * Sgn(AddtoUserPos.Y))
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = 0
                End If
            End If
    
            
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
            If ModoTrabajo Then Call Dialogos.DrawText(260, 260, "MODO TRABAJO", vbRed)
                If TaInvi > 20 Then Call Dialogos.DrawText(260, 275, "TIEMPO INVISIBLE " & Int(TaInvi / 30), vbWhite)

                If Cartel Then Call DibujarCartel
                
                If Dialogos.CantidadDialogos <> 0 Then Call Dialogos.MostrarTexto
                Call DrawBackBufferSurface
               
                Call RenderSounds
                
                
                
                
                
                
    
            
            
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
    End If
    
    
    
    If (GetTickCount - LastTime > 20) Then
        If Not Pausa And frmMain.Visible And Not frmForo.Visible Then
            CheckKeys
            LastTime = GetTickCount
        End If
    End If
    
    If Musica = 0 Then
        If Not SegState Is Nothing Then
            If Not Perf.IsPlaying(Seg, SegState) Then Play_Midi
        End If
    End If
         
    
    
    
    
    
        
        If DirectX.TickCount - lFrameTimer > 1000 Then
            FramesPerSec = FramesPerSecCounter
            If FPSFLAG Then frmMain.Caption = "Fenix AO" & " V " & App.Major & "." & App.Minor & "." & App.Revision
            frmMain.fpstext.Caption = FramesPerSec
            FramesPerSecCounter = 0
            lFrameTimer = DirectX.TickCount
        End If
        
        
        
        
        
            While DirectX.TickCount - lFrameLimiter < 55
                Sleep 5
            Wend
        
        
        

        lFrameLimiter = DirectX.TickCount
    
    
    
    
    esttick = GetTickCount
    For loopc = 1 To UBound(timers)
   
    
        timers(loopc) = timers(loopc) + (esttick - ulttick)
        
        If timers(1) >= tUs Then
            timers(1) = 0
            NoPuedeUsar = False
        End If
        
        
    Next loopc
    ulttick = GetTickCount
    DoEvents
Loop

EngineRun = False
frmCargando.Show
AddtoRichTextBox frmCargando.Status, "Liberando recursos...", 0, 0, 0, 0, 0, 1
LiberarObjetosDX

If bNoResChange = False Then
    Dim typDevM As typDevMODE
    Dim lRes As Long
    
    lRes = EnumDisplaySettings(0, 0, typDevM)
    With typDevM
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        .dmPelsWidth = oldResWidth
       .dmPelsHeight = oldResHeight
    End With
    lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
End If

Call UnloadAllForms

Config_Inicio.tip = tipf
Call EscribirGameIni(Config_Inicio)

End

ManejadorErrores:
    LogError "Contexto:" & Err.HelpContext & " Desc:" & Err.Description & " Fuente:" & Err.source
    End
    
End Sub



Sub WriteVar(File As String, Main As String, Var As String, value As String)




writeprivateprofilestring Main, Var, value, File

End Sub

Function GetVar(File As String, Main As String, Var As String) As String




Dim l As Integer
Dim Char As String
Dim sSpaces As String
Dim szReturn As String

szReturn = ""

sSpaces = Space(5000)


getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File

GetVar = RTrim(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function
Public Sub BMPtoGIF(bmp_fname As String, gif_fname As String)
Dim bdat As BITMAPINFOHEADER
Dim tmpimage As imgdes
Dim tmpimage2 As imgdes

Call BMPInfo(bmp_fname, bdat)
Call allocimage(tmpimage, bdat.biWidth, bdat.biHeight, 24)
Call loadbmp(bmp_fname, tmpimage)

Call allocimage(tmpimage2, bdat.biWidth, bdat.biHeight, 8)
Call convertrgbtopalex(256, tmpimage, tmpimage2, 3)

Call savegifex(gif_fname, tmpimage2, 8, 0)
Call Kill(bmp_fname)

End Sub
Public Function CheckMailString(ByRef sString As String) As Boolean
On Error GoTo errHnd:
Dim lPos As Long, lX As Long

lPos = InStr(sString, "@")
If (lPos <> 0) Then
    If Not InStr(lPos, sString, ".", vbBinaryCompare) > (lPos + 1) Then Exit Function

    For lX = 0 To Len(sString) - 1
        If Not lX = (lPos - 1) And Not CMSValidateChar_(Asc(Mid$(sString, (lX + 1), 1))) Then Exit Function
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
Function HayAgua(X As Integer, Y As Integer) As Boolean

If MapData(X, Y).Graphic(1).GrhIndex >= 1505 And _
   MapData(X, Y).Graphic(1).GrhIndex <= 1520 And _
   MapData(X, Y).Graphic(2).GrhIndex = 0 Then
            HayAgua = True
Else
            HayAgua = False
End If

End Function



    Public Sub ShowSendTxt()
        If Not frmCantidad.Visible Then
            frmMain.SendTxt.Visible = True
            frmMain.SendTxt.SetFocus
        End If
    End Sub
    

