Attribute VB_Name = "modTileEngine"
Option Explicit
'Techos con transparencias
Public VelocidadDesvanecimiento As Integer
Public ColorPromedio As Long
Public CDesvanecimiento(3) As Long
Public VDesvanecimiento As Integer
Public ControldeValorDesvanecimiento As Integer

'***************** vbGore Declares - Particles
Public ParticleOffsetX As Long
Public ParticleOffsetY As Long
Public LastOffsetX As Integer
Public LastOffsetY As Integer
Public LastTexture As Long
Public PixelOffsetX As Integer
Public PixelOffsetY As Integer
Public minY As Integer          'Start Y pos on current screen + tilebuffer
Public maxY As Integer          'End Y pos on current screen
Public minX As Integer          'Start X pos on current screen
Public maxX As Integer          'End X pos on current screen
Public ScreenMinY As Integer    'Start Y pos on current screen
Public ScreenMaxY As Integer    'End Y pos on current screen
Public ScreenMinX As Integer    'Start X pos on current screen
Public ScreenMaxX As Integer    'End X pos on current screen
Public PartMaxX As Integer
Public PartMaxY As Integer

Public Const ScreenWidth As Long = 541 'Keep this identical to the value on the server!
Public Const ScreenHeight As Long = 416 'Keep this identical to the value on the server!
Public ParticleTexture(1 To 12) As Direct3DTexture8
 
Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180
Public Const RadianToDegree As Single = 57.2958279087977 '180 / Pi
 
Public OffsetCounterX As Single
Public OffsetCounterY As Single
'***************** vbGore Declares - Particles





Public Const PI As Single = 3.14159265358979

Const HASH_TABLE_SIZE As Long = 337
Private Const BYTES_PER_MB As Long = 1048576                        '1Mb = 1024 Kb = 1024 * 1024 bytes = 1048576 bytes
Private Const MIN_MEMORY_TO_USE As Long = 16 * BYTES_PER_MB          '4 Mb

Private Type SURFACE_ENTRY_DYN
    FileName As Integer
    UltimoAcceso As Long
    Texture As Direct3DTexture8
    size As Long
    texture_width As Integer
    texture_height As Integer
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private TexList(HASH_TABLE_SIZE - 1) As HashNode

Private mD3D As D3DX8
Private device As Direct3DDevice8
 
Private mCantidadGraficos As Integer
Private maxBytesToUse As Long

Private lFrameLimiter As Long
Public lFrameModLimiter As Long
Public lFrameTimer As Long
Public timerTicksPerFrame As Single
Public timerElapsedTime As Single
Public particletimer As Single
Public engineBaseSpeed As Single

' Vector Usado para los Quads
Public Vector(3) As TLVERTEX
 
' INDEX BUFFERS
Public vbQuadIdx As DxVBLibA.Direct3DVertexBuffer8
Public ibQuad As DxVBLibA.Direct3DIndexBuffer8
Public indexList(0 To 5) As Integer 'the 6 indices required (note that the number is the
                              'same as the vertex count in the previous version).

'Describes a transformable lit vertex
Public Type TLVERTEX
  X As Single
  Y As Single
  Z As Single
  rhw As Single
  Color As Long
  Specular As Long
  tu As Single
  tv As Single
End Type

'********** Direct X ***********
Private Type D3D8Textures
    Texture As Direct3DTexture8
    texwidth As Integer
    texheight As Integer
End Type

'DirectX 8 Objects
Public Dx As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8
Public D3DDevice As Direct3DDevice8

'Font List
Public FontList As D3DXFont
Public FontDesc As IFont


Private Type light
    Active As Boolean 'Do we ignore this light?
    id As Long
    map_x As Integer 'Coordinates
    map_y As Integer
    Color As Long 'Start colour
    range As Byte
End Type

'Light list
Dim light_list() As light
Dim light_count As Long
Dim light_last As Long

Public CBlanco(3) As Long

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

'To get free bytes in drive
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, BytesTotal As Currency, FreeBytesTotal As Currency) As Long

Public mFreeMemoryBytes As Long

Private pUdtMemStatus As MEMORYSTATUS

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Dim valoreBlur As Long
Dim dimeTex As Long
Dim tex As Direct3DTexture8
Dim D3DbackBuffer As Direct3DSurface8
Dim zTarget As Direct3DSurface8
Dim stencil As Direct3DSurface8
Dim superTex As Direct3DSurface8
Dim blur As Boolean
Public blur_factor As Byte

Dim bump_map_texture As Direct3DTexture8
Dim bump_map_texture_ex As Direct3DTexture8
Dim bump_map_supported As Boolean
Dim bump_map_powa As Boolean

Public base_light As Long
Public day_r_old As Byte
Public day_g_old As Byte
Public day_b_old As Byte
Type luzxhora
    r As Long
    G As Long
    b As Long
End Type
Public luz_dia(0 To 24) As luzxhora

Public Const ImgSize As Byte = 4

Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

Public Const GrhFogata = 1521


Public Const SRCCOPY = &HCC0020

Public Type Position
    X As Integer
    Y As Integer
End Type

Public Type Position2
    X As Single
    Y As Single
End Type


Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames(1 To 25) As Integer
    Speed As Integer
    Active As Boolean
    
    MiniMap_color As Long
End Type
 
Public Type Grh
    Loops        As Integer
    GrhIndex     As Integer
    FrameCounter As Single
    SpeedCounter As Single
    Started      As Byte
    angle        As Single
End Type

Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

Public Type HeadData
    Head(1 To 4) As Grh
End Type

Type WeaponAnimData
    WeaponWalk(1 To 4) As Grh
End Type

Type ShieldAnimData
    ShieldWalk(1 To 4) As Grh
End Type

Public Type FxData
    FX As Grh
    OffsetX As Long
    OffsetY As Long
End Type

Public Type Char
    ParticleIndex As Integer
    Active As Byte
    Heading As Byte
    POS As Position

    Body As BodyData
    Head As HeadData
    casco As HeadData
    arma As WeaponAnimData
    escudo As ShieldAnimData
    UsandoArma As Boolean
    FX As Integer
    FxLoopTimes As Integer
    Criminal As Byte
    Navegando As Byte
    
    Nombre As String
    GM As Integer
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    haciendoataque As Byte
    Moving As Byte
    MoveOffset As Position2
    ServerIndex As Integer
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    
End Type

Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh

    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
    
    light_value(3) As Long
    
    luz As Integer
    Color(3) As Long
    
    ParticleIndex As Integer
End Type

Public IniPath As String
Public MapPath As String

Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

Public CurMap As Integer
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position
Public AddtoUserPos As Position
Public UserCharIndex As Integer

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long

Public WindowTileWidth As Integer
Public WindowTileHeight As Integer
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer
Public ScrollPixelsPerFrame As Single

Public LastChar As Integer

Public GrhData() As GrhData
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public Grh() As Grh
Public MapData() As MapBlock
Public CharList(1 To 10000) As Char
 
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
 

Public bRain        As Boolean
Public bTecho       As Boolean

Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
    plFogata = 3
End Enum

'Describes the return from a texture init
Private Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type
Private Type VFH
    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte 'The actual factual width of each character
    CharVA(0 To 255) As CharVA
End Type
Private Type CustomFont
    HeaderInfo As VFH           'Holds the header information
    Texture As Direct3DTexture8 'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As POINTAPI     'Size of the texture
End Type
 
'Private Const Font_Default_TextureNum As Long = -1   'The texture number used to represent this font - only used for AlternateRendering - keep negative to prevent interfering with game textures
Private cfonts() As CustomFont

Public HalfWindowTileWidth As Integer
Public HalfWindowTileHeight As Integer
Public TileBufferSize As Integer
Public TileBufferPixelOffsetX As Integer
Public TileBufferPixelOffsetY As Integer

Private Type FloatSurface
    POS As WorldPos
    offset As Position
    Grh As Grh
End Type

Public LastBlood As Integer     'Last blood splatter index used
Public BloodList() As FloatSurface

'BitBlt
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef source As Any, ByVal length As Long)

Public Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Public Function Engine_TPtoSPX(ByVal X As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'More info: http://www.vbgore.com/GameClient.TileEn ... ne_TPtoSPX" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'************************************************************
    Engine_TPtoSPX = X * 32 - ScreenMinX * 32 + OffsetCounterX - 16
End Function
 
Public Function Engine_TPtoSPY(ByVal Y As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'More info: http://www.vbgore.com/GameClient.TileEn ... ne_TPtoSPY" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'************************************************************
    Engine_TPtoSPY = Y * 32 - ScreenMinY * 32 + OffsetCounterY - 16
   
End Function


Function Engine_PixelPosX(ByVal X As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'More info: http://www.vbgore.com/GameClient.TileEn ... _PixelPosX" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
 
    Engine_PixelPosX = (X - 1) * TilePixelWidth
 
End Function
 
Function Engine_PixelPosY(ByVal Y As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'More info: http://www.vbgore.com/GameClient.TileEn ... _PixelPosY" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
 
    Engine_PixelPosY = (Y - 1) * TilePixelWidth
End Function



Public Function Engine_SPtoTPX(ByVal X As Long) As Long
 
'************************************************************
'Screen Position to Tile Position
'Takes the screen pixel position and returns the tile position
'************************************************************
 
    Engine_SPtoTPX = UserPos.X + X \ TilePixelWidth - WindowTileWidth \ 2
 
End Function
 
Public Function Engine_SPtoTPY(ByVal Y As Long) As Long
 
'************************************************************
'Screen Position to Tile Position
'Takes the screen pixel position and returns the tile position
'************************************************************
 
    Engine_SPtoTPY = UserPos.Y + Y \ TilePixelHeight - WindowTileHeight \ 2
 
End Function

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetAngle
'************************************************************
Dim SideA As Single
Dim SideC As Single

    On Error GoTo ErrOut

    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then

        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90

            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If

        'Exit the function
        Exit Function

    End If

    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then

        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360

            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If

        'Exit the function
        Exit Function

    End If

    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)

    'Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)

    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle

    'Exit function

Exit Function

    'Check for error
ErrOut:

    'Return a 0 saying there was an error
    Engine_GetAngle = 0

Exit Function

End Function

Public Sub Engine_Blood_Create(ByVal X As Integer, ByVal Y As Integer)
'*****************************************************************
'Create a blood splatter
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Blood_Create
'*****************************************************************
Dim BloodIndex As Integer

    'Get the next open blood slot
    Do
        BloodIndex = BloodIndex + 1

        'Update LastBlood if we go over the size of the current array
        If BloodIndex > LastBlood Then
            LastBlood = BloodIndex
            ReDim Preserve BloodList(1 To LastBlood)
            Exit Do
        End If

    Loop While BloodList(BloodIndex).Grh.GrhIndex > 0

    'Fill in the values
    BloodList(BloodIndex).POS.X = X
    BloodList(BloodIndex).POS.Y = Y
    InitGrh BloodList(BloodIndex).Grh, 21

End Sub

Public Sub Engine_Blood_Erase(ByVal BloodIndex As Integer)
'*****************************************************************
'Erase a blood splatter
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Blood_Erase
'*****************************************************************

    'Clear the selected index
    BloodList(BloodIndex).Grh.FrameCounter = 0
    BloodList(BloodIndex).Grh.GrhIndex = 0
    BloodList(BloodIndex).POS.X = 0
    BloodList(BloodIndex).POS.Y = 0

    'Update LastBlood
    If BloodIndex = LastBlood Then
        Do Until BloodList(LastBlood).Grh.GrhIndex > 1

            'Move down one splatter
            LastBlood = LastBlood - 1

            If LastBlood = 0 Then
                Erase BloodList
                Exit Sub
            Else
                'We still have blood, resize the array to end at the last used slot
                ReDim Preserve BloodList(1 To LastBlood)
            End If

        Loop
    End If

End Sub

Public Sub ShowNextFrame()


Dim ulttick As Long, esttick As Long
Dim timers(1 To 5) As Long
Dim loopc As Long

Do While prgRun
    
    If EngineRun Then
        If frmMain.WindowState <> 1 Then

            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrame * AddtoUserPos.X * timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
     
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
               OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrame * AddtoUserPos.Y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
            

    
    D3DDevice.BeginScene
     D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
            
            
            If UserCiego Then
                D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
            Else
                RenderScreen UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY
            End If
            
            If ModoTrabajo Then DrawText 1, 10, 10, "MODO TRABAJO", D3DColorXRGB(255, 0, 0)
            If Cartel Then DibujarCartel
            Dialogos.Render
            RenderSounds
                
        D3DDevice.Present ByVal 0, ByVal 0, frmMain.Renderer.hWnd, ByVal 0
    D3DDevice.EndScene
    

    If frmMain.Inventario.Visible Then
        DrawInventario
        End If
            lFrameLimiter = GetTickCount
            FramesPerSecCounter = FramesPerSecCounter + 1
            timerElapsedTime = GetElapsedTime()
            timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
            particletimer = timerElapsedTime * 0.05
        End If
    End If
    
    If Not Pausa And frmMain.Visible And Not frmForo.Visible Then
        CheckKeys
    End If

    If GetTickCount - lFrameTimer > 1000 Then
        FramesPerSec = FramesPerSecCounter
        If FPSFLAG Then frmMain.Caption = "Fenix AO" & " V " & App.Major & "." & App.Minor & "." & App.Revision
        frmMain.fpstext.Caption = FramesPerSec
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount
    End If
    
    
    ' ### I N T E R V A L O S ###
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

End Sub

Sub DrawInventario()
If ActualizarInv = False Then Exit Sub

    Dim re As RECT
    re.left = 0
    re.top = 0
    re.bottom = 176
    re.Right = 160
   
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene
   
    Dim i As Byte, X As Integer, Y As Integer
    Dim T As Grh
 
    For Y = 1 To 5
        For X = 1 To 5
        i = i + 1
       
        If UserInventory(i).GrhIndex Then
           
            InitGrh T, UserInventory(i).GrhIndex
           
            With UserInventory(i)
            
                Device_Box_Textured_Render GrhData(UserInventory(i).GrhIndex).FileNum, X * 32 - 32, Y * 32 - 32, GrhData(UserInventory(i).GrhIndex).pixelWidth, GrhData(UserInventory(i).GrhIndex).pixelHeight, CBlanco(), 0, 0
                If ItemElegido = i Then Device_Box_Textured_Render 11000, X * 32 - 32, Y * 32 - 32, 32, 32, CBlanco(), 0, 0
                
                DrawText 2, X * 32 - 32, Y * 32 - 32 - 2, UserInventory(i).Amount, D3DColorARGB(255, 255, 255, 255)
                
               
                If UserInventory(i).Equipped Then _
                DrawText 2, (X * 32) + 22 - 32, (Y * 32) + 20 - 32 - 1, "+", D3DColorARGB(255, 255, 255, 0)
                
                
            End With
        End If
       
 
    Next X, Y
 
    D3DDevice.EndScene
    D3DDevice.Present re, ByVal 0, frmMain.Inventario.hWnd, ByVal 0
    
    ActualizarInv = False
 
End Sub

Sub Draw_Grh(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, ByRef Color() As Long, Optional Alpha As Boolean, Optional ByVal shadow As Byte = 0, Optional ByVal Invert_x As Boolean = False, Optional ByVal Invert_y As Boolean = False, Optional ByVal KillAnim As Integer = 0, Optional ByVal map_x As Byte, Optional ByVal map_y As Byte)
On Error Resume Next
Dim iGrhIndex As Integer
Dim QuitarAnimacion As Boolean


If Animate Then
    If Grh.Started = 1 Then
       
        Grh.FrameCounter = Grh.FrameCounter + ((timerElapsedTime * 0.1) * GrhData(Grh.GrhIndex).NumFrames / Grh.SpeedCounter)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
               
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                   
                If KillAnim <> 0 Then
                If CharList(KillAnim).FX > 0 Then
                    If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                          CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes <= 0 Then CharList(KillAnim).FX = 0: Exit Sub
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

If map_x Or map_y = 0 Then map_x = 1: map_y = 1

Call Device_Box_Textured_Render(GrhData(iGrhIndex).FileNum, _
        X, Y, _
        GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
        Color(), _
        GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
        Alpha, Grh.angle)

End Sub

Sub DrawGrhtoHdc(hDC As Long, GrhIndex As Integer)

    Dim hDCsrc As Long
 
    If GrhIndex <= 0 Then Exit Sub
        
        'If it's animated switch GrhIndex to first frame
        If GrhData(GrhIndex).NumFrames <> 1 Then
            GrhIndex = GrhData(GrhIndex).Frames(1)
        End If
           
        hDCsrc = CreateCompatibleDC(hDC)
        
        Call SelectObject(hDCsrc, LoadPicture(App.Path & "\Graficos\" & GrhData(GrhIndex).FileNum & ".bmp"))

        'Draw
        BitBlt hDC, 0, 0, _
        GrhData(GrhIndex).pixelWidth, GrhData(GrhIndex).pixelHeight, _
        hDCsrc, _
        GrhData(GrhIndex).sX, GrhData(GrhIndex).sY, _
        vbSrcCopy

        DeleteDC hDCsrc
End Sub

Public Sub Dibujar_grh_Simple(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, Optional Color As Long)
Dim C(3) As Long
 
If Grh.GrhIndex = 0 Then Exit Sub
 
C(0) = Color
C(1) = Color
C(2) = Color
C(3) = Color
 
If Grh.FrameCounter = 0 Then Grh.FrameCounter = 2
 
With GrhData(Grh.GrhIndex)
 
    Device_Box_Textured_Render Grh.GrhIndex, X, Y, .pixelWidth, .pixelHeight, C(), .sX, .sY
 
End With
 
End Sub

Public Sub Draw_FilledBox(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, Color As Long, outlinecolor As Long)
 
    Static box_rect As RECT
    Static Outline As RECT
    Static rgb_list(3) As Long
    Static rgb_list2(3) As Long
    Static Vertex(3) As TLVERTEX
    Static Vertex2(3) As TLVERTEX
   
    rgb_list(0) = Color
    rgb_list(1) = Color
    rgb_list(2) = Color
    rgb_list(3) = Color
   
    rgb_list2(0) = outlinecolor
    rgb_list2(1) = outlinecolor
    rgb_list2(2) = outlinecolor
    rgb_list2(3) = outlinecolor
   
    With box_rect
        .bottom = Y + Height - 1
        .left = X + 1
        .Right = X + Width - 1
        .top = Y + 1
    End With
   
    With Outline
        .bottom = Y + Height
        .left = X
        .Right = X + Width
        .top = Y
    End With
   
   
    Geometry_Create_Box Vertex2(), Outline, Outline, rgb_list2(), 0, 0
    Geometry_Create_Box Vertex(), box_rect, box_rect, rgb_list(), 0, 0
   
   
    D3DDevice.SetTexture 0, Nothing
    'D3DDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLESTRIP, 0, 4, 2, indexList(0), D3DFMT_INDEX16, Vertex2(0), Len(Vertex2(0))
    'D3DDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLESTRIP, 0, 4, 2, indexList(0), D3DFMT_INDEX16, Vertex(0), Len(Vertex(0))
 
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex2(0), Len(Vertex2(0))
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
   
End Sub

Sub RenderScreen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim Y                   As Integer     'Keeps track of where on map we are
    Dim X                   As Integer     'Keeps track of where on map we are
    Dim ScreenX             As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY             As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset          As Integer
    Dim minYOffset          As Integer
    Dim PixelOffsetXTemp    As Integer 'For centering grhs
    Dim PixelOffsetYTemp    As Integer 'For centering grhs
    Dim CurrentGrhIndex     As Integer
    Dim offx                As Integer
    Dim offy                As Integer
    Dim TempChar As Char
    Dim Moved    As Byte
    Dim iPPx     As Integer
    Dim iPPy     As Integer

    'Figure out Ends and Starts of screen
    ScreenMinY = TileY - HalfWindowTileHeight
    ScreenMaxY = TileY + HalfWindowTileHeight
    ScreenMinX = TileX - HalfWindowTileWidth
    ScreenMaxX = TileX + HalfWindowTileWidth
    
    minY = ScreenMinY - TileBufferSize
    maxY = ScreenMaxY + TileBufferSize
    minX = ScreenMinX - TileBufferSize
    maxX = ScreenMaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If ScreenMinY > YMinMapSize Then
        ScreenMinY = ScreenMinY - 1
    Else
        ScreenMinY = 1
        ScreenY = 1
    End If
    
    If ScreenMaxY < YMaxMapSize Then ScreenMaxY = ScreenMaxY + 1
    
    If ScreenMinX > XMinMapSize Then
        ScreenMinX = ScreenMinX - 1
    Else
        ScreenMinX = 1
        ScreenX = 1
    End If
    
    If ScreenMaxX < XMaxMapSize Then ScreenMaxX = ScreenMaxX + 1

    ParticleOffsetX = (Engine_PixelPosX(ScreenMinX) - PixelOffsetX)
    ParticleOffsetY = (Engine_PixelPosY(ScreenMinY) - PixelOffsetY)
    

    
    'Draw floor layer
    For Y = ScreenMinY To ScreenMaxY
        For X = ScreenMinX To ScreenMaxX
            'Layer 1 **********************************

                Call Draw_Grh(MapData(X, Y).Graphic(1), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 0, 1, MapData(X, Y).light_value(), , , , , , X, Y)
                
                If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call Draw_Grh(MapData(X, Y).Graphic(2), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 1, 1, MapData(X, Y).light_value(), , , , , , X, Y)
                End If
            '******************************************
            ScreenX = ScreenX + 1
        Next X

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + ScreenMinX
        ScreenY = ScreenY + 1
    Next Y
    
    
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            With MapData(X, Y)
                '******************************************

                'Object Layer **********************************
                If .ObjGrh.GrhIndex <> 0 Then

                    Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value(), , , , , , X, Y)
                End If


                If .CharIndex <> 0 Then
                                    TempChar = CharList(MapData(X, Y).CharIndex)
                PixelOffsetXTemp = PixelOffsetX
                PixelOffsetYTemp = PixelOffsetY
                Moved = 0

            With TempChar
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffset.X = .MoveOffset.X + ScrollPixelsPerFrame * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).SpeedCounter > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .arma.WeaponWalk(.Heading).Started = 1
                .escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                Moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffset.X >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffset.X <= 0) Then
                    .MoveOffset.X = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffset.Y = .MoveOffset.Y + ScrollPixelsPerFrame * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).SpeedCounter > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .arma.WeaponWalk(.Heading).Started = 1
                .escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                Moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffset.Y >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffset.Y <= 0) Then
                    .MoveOffset.Y = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
 
            If .Heading = 0 Then .Heading = 3
 
            If Moved = 0 Then
                .Body.Walk(.Heading).Started = 0
                .Body.Walk(.Heading).FrameCounter = 1
               
                .arma.WeaponWalk(.Heading).Started = 0
                .arma.WeaponWalk(.Heading).FrameCounter = 1
               
                .escudo.ShieldWalk(.Heading).Started = 0
                .escudo.ShieldWalk(.Heading).FrameCounter = 1
               
                .Moving = 0
            End If
           
            If TempChar.haciendoataque = 0 And .MoveOffset.X = 0 And .MoveOffset.Y = 0 Then
                .arma.WeaponWalk(.Heading).Started = 0
                '.arma.WeaponWalk(.Heading).FrameCounter = 1
                .escudo.ShieldWalk(.Heading).Started = 0
               
                End If
               
            If TempChar.haciendoataque = 1 Then
                .arma.WeaponWalk(.Heading).Started = 1
                .escudo.ShieldWalk(.Heading).Started = 1
                .haciendoataque = 0
            End If
           
    End With
    PixelOffsetXTemp = PixelOffsetXTemp + TempChar.MoveOffset.X
    PixelOffsetYTemp = PixelOffsetYTemp + TempChar.MoveOffset.Y
    
                iPPx = ((32 * ScreenX) - 32) + PixelOffsetXTemp + 32
                iPPy = ((32 * ScreenY) - 32) + PixelOffsetYTemp + 32
                
                If Len(TempChar.Nombre) = 0 Then
                        Call Draw_Grh(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value())
                        'Cabeza
                        If TempChar.Head.Head(TempChar.Heading).GrhIndex > 0 Then
                        Call Draw_Grh(TempChar.Head.Head(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y, 1, 0, MapData(X, Y).light_value())
                        End If
                Else
                    If TempChar.Navegando = 1 Then
                        'Cuerpo (Barca / Galeon / Galera)
                        Call Draw_Grh(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value())
                    
                    ElseIf Not CharList(MapData(X, Y).CharIndex).invisible And TempChar.Head.Head(TempChar.Heading).GrhIndex > 0 Then
                        
                        'Cuerpo
                        Call Draw_Grh(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value)
                        
                        If UserParalizado Then
                            Call DDrawTono(TempChar.Body.Walk(TempChar.Heading), iPPx + 1, iPPy - 12, 0, 0, 0, 0, 0, 0, 0, 0)
                        End If
                        
                        'Cabeza
                        If TempChar.Head.Head(TempChar.Heading).GrhIndex > 0 Then
                            Call Draw_Grh(TempChar.Head.Head(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y, 1, 0, MapData(X, Y).light_value())
                            If UserParalizado Then
                                Call DDrawTono(TempChar.Head.Head(TempChar.Heading), iPPx + 8, iPPy - 22, 0, 0, 0, 0, 0, 0, 0, 0)
                            End If
                        End If
                        
                        'Casco
                        If TempChar.casco.Head(TempChar.Heading).GrhIndex > 0 Then
                            Call Draw_Grh(TempChar.casco.Head(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y, 1, 0, MapData(X, Y).light_value())
                        End If
                        
                        'Arma
                        If TempChar.arma.WeaponWalk(TempChar.Heading).GrhIndex > 0 Then
                            Call Draw_Grh(TempChar.arma.WeaponWalk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value())
                        End If
                        
                        'Escudo
                        If TempChar.escudo.ShieldWalk(TempChar.Heading).GrhIndex > 0 Then
                            Call Draw_Grh(TempChar.escudo.ShieldWalk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value())
                        End If

                    ElseIf (TempChar.invisible And (CharList(UserCharIndex).Nombre = CharList(MapData(X, Y).CharIndex).Nombre)) Or PJInMyGuild(MapData(X, Y).CharIndex) Then

                        'Cuerpo
                        Call Draw_Grh(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value, True)
    
                        'Cabeza
                        If TempChar.Head.Head(TempChar.Heading).GrhIndex > 0 Then
                            Call Draw_Grh(TempChar.Head.Head(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y, 1, 0, MapData(X, Y).light_value(), True)
                        End If
    
                        'Casco
                        If TempChar.casco.Head(TempChar.Heading).GrhIndex > 0 Then
                            Call Draw_Grh(TempChar.casco.Head(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y, 1, 0, MapData(X, Y).light_value(), True)
                        End If
    
                        'Arma
                        If TempChar.arma.WeaponWalk(TempChar.Heading).GrhIndex > 0 Then
                            Call Draw_Grh(TempChar.arma.WeaponWalk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value(), True)
                        End If
    
                        'Escudo
                        If TempChar.escudo.ShieldWalk(TempChar.Heading).GrhIndex > 0 Then
                            Call Draw_Grh(TempChar.escudo.ShieldWalk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value(), True)
                        End If

                    End If
                        
                    If Nombres Then
                        Dim sClan As String
                        If Not (TempChar.invisible Or TempChar.Navegando = 1) Then
                            Dim lCenter As Long
                            If InStr(TempChar.Nombre, "<") > 0 And InStr(TempChar.Nombre, ">") > 0 Then
                                lCenter = (frmMain.textwidth(left$(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1)) / 2) - 16
                                sClan = mid$(TempChar.Nombre, InStr(TempChar.Nombre, "<"))
                                Call DrawText(1, iPPx - lCenter, iPPy + 30, left$(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), D3DColorXRGB(RG(TempChar.Criminal, 1), RG(TempChar.Criminal, 2), RG(TempChar.Criminal, 3)))
                                lCenter = (frmMain.textwidth(sClan) / 2) - 16
                                Call DrawText(1, iPPx - lCenter, iPPy + 45, sClan, D3DColorXRGB(RG(TempChar.Criminal, 1), RG(TempChar.Criminal, 2), RG(TempChar.Criminal, 3)))
                            Else
                                lCenter = (frmMain.textwidth(TempChar.Nombre) / 2) - 16
                                Call DrawText(1, iPPx - lCenter, iPPy + 30, TempChar.Nombre, D3DColorXRGB(RG(TempChar.Criminal, 1), RG(TempChar.Criminal, 2), RG(TempChar.Criminal, 3)))
                            End If
                        ElseIf (TempChar.invisible And (CharList(UserCharIndex).Nombre = CharList(MapData(X, Y).CharIndex).Nombre)) Or PJInMyGuild(MapData(X, Y).CharIndex) Then
                            If InStr(TempChar.Nombre, "<") > 0 And InStr(TempChar.Nombre, ">") > 0 Then
                                lCenter = (frmMain.textwidth(left$(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1)) / 2) - 16
                                sClan = mid$(TempChar.Nombre, InStr(TempChar.Nombre, "<"))
                                Call DrawText(1, iPPx - lCenter, iPPy + 30, left$(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), D3DColorXRGB(255, 255, 0))
                                lCenter = (frmMain.textwidth(sClan) / 2) - 16
                                Call DrawText(1, iPPx - lCenter, iPPy + 45, sClan, D3DColorXRGB(255, 255, 0))
                            Else
                                lCenter = (frmMain.textwidth(TempChar.Nombre) / 2) - 16
                                Call DrawText(1, iPPx - lCenter, iPPy + 30, TempChar.Nombre, D3DColorXRGB(255, 255, 0))
                            End If
                        End If
                    End If
                End If
    
                Call Dialogos.UpdateDialogPos((iPPx + TempChar.Body.HeadOffset.X), (iPPy + TempChar.Body.HeadOffset.Y), MapData(X, Y).CharIndex)
                
                CharList(MapData(X, Y).CharIndex) = TempChar

                If CharList(MapData(X, Y).CharIndex).FX <> 0 Then Call Draw_Grh(FxData(TempChar.FX).FX, iPPx + FxData(TempChar.FX).OffsetX, iPPy + FxData(TempChar.FX).OffsetY, 1, 1, CBlanco(), , , , , MapData(X, Y).CharIndex)
                
            End If
                '*************************************************
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then

                    Call Draw_Grh(.Graphic(3), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, MapData(X, Y).light_value(), , , , , , X, Y)
                End If
                '************************************************

            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    ScreenY = minYOffset - 5

     'ParticleOffsetX = (Engine_PixelPosX(ScreenMinX) - PixelOffsetX)
    'ParticleOffsetY = (Engine_PixelPosY(ScreenMinY) - PixelOffsetY)
 
Effect_UpdateAll

 
If Not bTecho Then
    'Draw blocked tiles and grid
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            'Layer 4 **********************************
            If MapData(X, Y).Graphic(4).GrhIndex Then
            
                If VDesvanecimiento < ColorPromedio Then
                    ControldeValorDesvanecimiento = ControldeValorDesvanecimiento + 1
                    If ControldeValorDesvanecimiento = VelocidadDesvanecimiento Then
                        VDesvanecimiento = VDesvanecimiento + 1
                        ControldeValorDesvanecimiento = 0
                    End If
                End If
            
            CDesvanecimiento(0) = D3DColorXRGB(VDesvanecimiento, VDesvanecimiento, VDesvanecimiento)
            CDesvanecimiento(1) = D3DColorXRGB(VDesvanecimiento, VDesvanecimiento, VDesvanecimiento)
            CDesvanecimiento(2) = D3DColorXRGB(VDesvanecimiento, VDesvanecimiento, VDesvanecimiento)
            CDesvanecimiento(3) = D3DColorXRGB(VDesvanecimiento, VDesvanecimiento, VDesvanecimiento)
    
            Call Draw_Techos(MapData(X, Y).Graphic(4), _
            ScreenX * 32 + PixelOffsetX, _
            ScreenY * 32 + PixelOffsetY, _
            1, 1)
            
            End If
            '**********************************
            ScreenX = ScreenX + 1
        Next X
        
        ScreenY = ScreenY + 1
    Next Y
Else
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
                'Layer 4 **********************************
                If MapData(X, Y).Graphic(4).GrhIndex Then
                
                If VDesvanecimiento > 0 Then
                ControldeValorDesvanecimiento = ControldeValorDesvanecimiento + 1
                If ControldeValorDesvanecimiento = VelocidadDesvanecimiento Then
                VDesvanecimiento = VDesvanecimiento - 1
                ControldeValorDesvanecimiento = 0
                End If
                End If
                
                CDesvanecimiento(0) = D3DColorXRGB(VDesvanecimiento, VDesvanecimiento, VDesvanecimiento)
                CDesvanecimiento(1) = D3DColorXRGB(VDesvanecimiento, VDesvanecimiento, VDesvanecimiento)
                CDesvanecimiento(2) = D3DColorXRGB(VDesvanecimiento, VDesvanecimiento, VDesvanecimiento)
                CDesvanecimiento(3) = D3DColorXRGB(VDesvanecimiento, VDesvanecimiento, VDesvanecimiento)
                
                Call Draw_Techos(MapData(X, Y).Graphic(4), _
                ScreenX * 32 + PixelOffsetX, _
                ScreenY * 32 + PixelOffsetY, _
                1, 1)
                End If
                '**********************************
               
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If

LastOffsetX = ParticleOffsetX
LastOffsetY = ParticleOffsetY

End Sub

Sub Draw_Techos(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0, Optional ByVal map_x As Byte, Optional ByVal map_y As Byte)
 
Dim iGrhIndex As Integer
Dim QuitarAnimacion As Boolean
 
 
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
 
If map_x Or map_y = 0 Then map_x = 1: map_y = 1
 
    Device_Box_Textured_Render_Advance iGrhIndex, _
        X, Y, _
        GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
        CDesvanecimiento(), _
        GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
        True, 0
 
End Sub

Sub Base_Luz(Rojo As Byte, Verde As Byte, Azul As Byte)
    base_light = D3DColorXRGB(Rojo, Verde, Azul)
    ColorPromedio = (Val(Rojo) + Val(Verde) + Val(Azul)) / 3
    Light_Render_All
End Sub

Public Function PJInMyGuild(ByVal CharI As Integer) As Boolean
    With CharList(CharI)
        Dim tp As Integer
        Dim tempS As String
        Dim tempP As Integer
        Dim MiTag As String
        Dim MiPos As Integer

        tp = InStr(.Nombre, "<")
        If tp = 0 Then
            PJInMyGuild = False
            Exit Function
        End If

        tempP = InStr(.Nombre, "<")
        MiPos = InStr(CharList(UserCharIndex).Nombre, "<")
        
        If MiPos = 0 Then
            PJInMyGuild = False
            Exit Function
        End If
        
        tempS = mid$(.Nombre, tempP)
        MiTag = mid$(CharList(UserCharIndex).Nombre, MiPos)
        If (tempS = MiTag) Then
            PJInMyGuild = True
            Exit Function
        End If
        PJInMyGuild = False
    End With
End Function

Public Function RenderSounds()

    If bLluvia(UserMap) = 1 Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> plLluviain Then
                    Call Audio.StopWave
                    Call Audio.PlayWave("lluviain.wav", 0, 0, Enabled)
                    frmMain.IsPlaying = plLluviain
                End If
                
                
            Else
                If frmMain.IsPlaying <> plLluviaout Then
                    Call Audio.StopWave
                    Call Audio.PlayWave("lluviaout.wav", 0, 0, Enabled)
                    frmMain.IsPlaying = plLluviaout
                End If
                
                
            End If
        End If
    End If

End Function

Public Sub CargarColores()
CBlanco(0) = D3DColorARGB(255, 255, 255, 255)
CBlanco(1) = D3DColorARGB(255, 255, 255, 255)
CBlanco(2) = D3DColorARGB(255, 255, 255, 255)
CBlanco(3) = D3DColorARGB(255, 255, 255, 255)
End Sub

Public Function General_Bytes_To_Megabytes(Bytes As Double) As Double
Dim dblAns As Double
dblAns = (Bytes / 1024) / 1024
General_Bytes_To_Megabytes = Format(dblAns, "###,###,##0.00")
End Function

Public Function General_Get_Free_Ram() As Double
    'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPhys
    General_Get_Free_Ram = General_Bytes_To_Megabytes(dblAns)
End Function

Public Function General_Get_Free_Ram_Bytes() As Long
    GlobalMemoryStatus pUdtMemStatus
    General_Get_Free_Ram_Bytes = pUdtMemStatus.dwAvailPhys
End Function

Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer) As Boolean

IniPath = App.Path & "\Init\"

UserPos.X = MinXBorder
UserPos.Y = MinYBorder

TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth

HalfWindowTileHeight = WindowTileHeight / 2
HalfWindowTileWidth = WindowTileWidth / 2

TileBufferSize = 9
TileBufferPixelOffsetX = (TileBufferSize - 1) * TilePixelWidth
TileBufferPixelOffsetY = (TileBufferSize - 1) * TilePixelHeight

MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

Call LoadGrhData
Call CargarCuerpos
Call CargarCabezas
Call CargarCascos
Call CargarFxs
Call CargarColores
Call CargarAnimArmas
Call CargarAnimEscudos
Call CargarAnimsExtra
Call CargarArrayLluvia
Call CargarMensajes
Call EstablecerRecompensas

'/////TSG: INICIAR DIRECT3D/////
Set Dx = New DirectX8
Set D3D = Dx.Direct3DCreate
Set D3DX = New D3DX8

Dim D3DWindow As D3DPRESENT_PARAMETERS
Dim DispMode As D3DDISPLAYMODE
Dim D3DCreate As CONST_D3DCREATEFLAGS

    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

With D3DWindow
        .Windowed = True
        Select Case GetVar(App.Path & "\Init\Opciones.opc", "CONFIG", "VSYNC")
        Case Is = 0
        .SwapEffect = D3DSWAPEFFECT_COPY
        Case Is = 1
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        End Select
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = frmMain.Renderer.ScaleWidth
        .BackBufferHeight = frmMain.Renderer.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.Renderer.hWnd
    End With
    DispMode.Format = D3DFMT_X8R8G8B8

DispMode.Format = D3DFMT_X8R8G8B8

Select Case GetVar(App.Path & "\Init\Opciones.opc", "CONFIG", "Iniciar")
Case "Mixed"
D3DCreate = D3DCREATE_MIXED_VERTEXPROCESSING

Case "Software"
D3DCreate = D3DCREATE_SOFTWARE_VERTEXPROCESSING

Case "Hardware"
D3DCreate = D3DCREATE_HARDWARE_VERTEXPROCESSING
End Select


Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.Renderer.hWnd, D3DCreate, _
                                                            D3DWindow)
    
    
    
    frmMain.Visible = False
    DoEvents
    
    D3DDevice.SetVertexShader FVF
    
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True

    mCantidadGraficos = 0
   
    'Seteamos el objeto
    Set mD3D = D3DX
    Set device = D3DDevice
    mFreeMemoryBytes = 0
    maxBytesToUse = MIN_MEMORY_TO_USE
    
    engineBaseSpeed = 0.017
    ScrollPixelsPerFrame = 9
    
'Load Index List
    indexList(0) = 0: indexList(1) = 1: indexList(2) = 2
    indexList(3) = 3: indexList(4) = 4: indexList(5) = 5
 
    Set ibQuad = D3DDevice.CreateIndexBuffer(Len(indexList(0)) * 4, 0, D3DFMT_INDEX16, D3DPOOL_MANAGED)
   
    D3DIndexBuffer8SetData ibQuad, 0, Len(indexList(0)) * 4, 0, indexList(0)
 
    ' Index Quad
    Set vbQuadIdx = D3DDevice.CreateVertexBuffer(Len(Vector(0)) * 4, 0, D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR, D3DPOOL_MANAGED)
    
    'partículas
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
    ' Initialize particles
    Call Engine_Init_ParticleEngine
    Call Base_Luz(255, 240, 180)

    ReDim cfonts(1 To Val(GetVar(App.Path & "\Init\Fuentes\Fuentes.dat", "INIT", "MaxFuentes"))) As CustomFont
    Engine_Init_FontTextures
    Engine_Init_FontSettings
    
    '/////TERMINA CARGA DE DIRECTX8/////
    
    VelocidadDesvanecimiento = 130
    InitTileEngine = True
End Function

Public Sub DeInitTileEngine()

    Dim i As Long
    Dim j As Long
    
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                Set .SurfaceEntry(j).Texture = Nothing
            Next j
            
            'Destroy the arrays
            Erase .SurfaceEntry
        End With
    Next i

    Set Dx = Nothing
    Set D3D = Nothing
    Set D3DX = Nothing
    Set D3DDevice = Nothing
    Set FontList = Nothing
    
    Dim loopc As Long
   
        'Clear particles
    For loopc = 1 To UBound(ParticleTexture)
        If Not ParticleTexture(loopc) Is Nothing Then Set ParticleTexture(loopc) = Nothing
    Next loopc
    
    Erase CharList
    Erase Grh
    Erase GrhData
    Erase MapData
End Sub

Private Function Engine_FToDW(f As Single) As Long
' single > long
Dim buf As D3DXBuffer
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, Engine_FToDW
End Function

Private Function VectorToRGBA(Vec As D3DVECTOR, fHeight As Single) As Long
Dim r As Integer, G As Integer, b As Integer, a As Integer
    r = 127 * Vec.X + 128
    G = 127 * Vec.Y + 128
    b = 127 * Vec.Z + 128
    a = 255 * fHeight
    VectorToRGBA = D3DColorARGB(a, r, G, b)
End Function

Public Function Light_Color_Value_Get(ByVal light_index As Long, ByRef color_value As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        color_value = light_list(light_index).Color
        Light_Color_Value_Get = True
    End If
End Function
Private Function Light_Check(ByVal light_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check light_index
    If light_index > 0 And light_index <= light_last Then
        If light_list(light_index).Active Then
            Light_Check = True
        End If
    End If
End Function
Public Function Light_Create(ByVal map_x As Integer, ByVal map_y As Integer, Optional ByVal color_value As Long = &HFFFFFFFF, _
                            Optional ByVal range As Byte = 1, Optional ByVal id As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
    If InMapBounds(map_x, map_y) Then
        'Make sure there is no light in the given map pos
        'If Map_Light_Get(map_x, map_y) <> 0 Then
        '    Light_Create = 0
        '    Exit Function
        'End If
        Light_Create = Light_Next_Open
        Light_Make Light_Create, map_x, map_y, color_value, range, id
    End If
End Function
Private Sub Light_Make(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, ByVal rgb_value As Long, _
                        ByVal range As Long, Optional ByVal id As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count + 1
    
    'Make active
    light_list(light_index).Active = True
    
    light_list(light_index).map_x = map_x
    light_list(light_index).map_y = map_y
    light_list(light_index).Color = rgb_value
    light_list(light_index).range = range
    light_list(light_index).id = id
End Sub
Public Sub Light_Render_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
            
    For loop_counter = 1 To light_count
        
        If light_list(loop_counter).Active Then
            Light_Render loop_counter
        End If
    
    Next loop_counter
End Sub

Private Sub Light_Render(ByVal light_index As Long)
'menduz
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim ia As Single
    Dim i As Integer
    Dim Color As Long
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Set color
    Color = light_list(light_index).Color
    
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(0) = Color
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(1) = Color
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(2) = Color
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(3) = Color
                
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = Color
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = Color
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = Color
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = Color
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).light_value(0) = Color
            MapData(X, min_y).light_value(2) = Color
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = Color
            MapData(X, max_y).light_value(3) = Color
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = Color
            MapData(min_x, Y).light_value(3) = Color
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = Color
            MapData(max_x, Y).light_value(1) = Color
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).light_value(0) = Color
                MapData(X, Y).light_value(1) = Color
                MapData(X, Y).light_value(2) = Color
                MapData(X, Y).light_value(3) = Color
            End If
        Next Y
    Next X
End Sub
Private Function Light_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).Active = False
        If loopc = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Next_Open = loopc
Exit Function
ErrorHandler:
    Light_Next_Open = 1
End Function

Public Function Light_Remove(ByVal light_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        Light_Destroy light_index
        Light_Remove = True
    End If
End Function

Public Function Light_Move(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then
        
            'Move it
            Light_Erase light_index
            light_list(light_index).map_x = map_x
            light_list(light_index).map_y = map_y
    
            Light_Move = True
            
        End If
    End If
End Function

Public Function Light_Move_By_Head(ByVal light_index As Long, ByVal Heading As Byte) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 15/05/2002
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Integer
    Dim map_y As Integer
    Dim nX As Integer
    Dim nY As Integer
    Dim addY As Byte
    Dim addX As Byte
    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Light_Move_By_Head = False
        Exit Function
    End If

    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
    
        map_x = light_list(light_index).map_x
        map_y = light_list(light_index).map_y
        


        Select Case Heading
            Case NORTH
                addY = -1
        
            Case EAST
                addX = 1
        
            Case SOUTH
                addY = 1
            
            Case WEST
                addX = -1
        End Select
        
        nX = map_x + addX
        nY = map_y + addY
        
        'Make sure it's a legal move
        If InMapBounds(nX, nY) Then
        
            'Move it
            Light_Erase light_index

            light_list(light_index).map_x = nX
            light_list(light_index).map_y = nY
    
            Light_Move_By_Head = True
            
        End If
    End If
End Function

Public Function Light_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).id = id
        If loopc = light_last Then
            Light_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Find = loopc
Exit Function
ErrorHandler:
    Light_Find = 0
End Function

Public Function Light_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To light_last
        'Make sure it's a legal index
        If Light_Check(Index) Then
            Light_Destroy Index
        End If
    Next Index
    
    Light_Remove_All = True
End Function

Private Sub Light_Destroy(ByVal light_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As light
    
    
    Light_Erase light_index
    
    light_list(light_index) = temp
    
    'Update array size
    If light_index = light_last Then
        Do Until light_list(light_last).Active
            light_last = light_last - 1
            If light_last = 0 Then
                light_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count - 1
End Sub

Private Sub Light_Erase(ByVal light_index As Long)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 3/31/2003
'Correctly erases a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer

    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = 0
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = 0
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = 0
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = 0
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).light_value(0) = 0
            MapData(X, min_y).light_value(2) = 0
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = 0
            MapData(X, max_y).light_value(3) = 0
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = 0
            MapData(min_x, Y).light_value(3) = 0
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = 0
            MapData(max_x, Y).light_value(1) = 0
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).light_value(0) = 0
                MapData(X, Y).light_value(1) = 0
                MapData(X, Y).light_value(2) = 0
                MapData(X, Y).light_value(3) = 0
            End If
        Next Y
    Next X
End Sub

Private Function CreateColorVal(a As Integer, r As Integer, G As Integer, b As Integer) As D3DCOLORVALUE
    CreateColorVal.a = a
    CreateColorVal.r = r
    CreateColorVal.G = G
    CreateColorVal.b = b
End Function
Public Function ARGB(ByVal r As Long, ByVal G As Long, ByVal b As Long, ByVal a As Long) As Long
        
    Dim C As Long
        
    If a > 127 Then
        a = a - 128
        C = a * 2 ^ 24 Or &H80000000
        C = C Or r * 2 ^ 16
        C = C Or G * 2 ^ 8
        C = C Or b
    Else
        C = a * 2 ^ 24
        C = C Or r * 2 ^ 16
        C = C Or G * 2 ^ 8
        C = C Or b
    End If
    
    ARGB = C

End Function

Public Function RGBtoD3DColorARGB(Alpha As Integer, ByVal Color As Long) As Long
    
Dim Rojo As Integer
Dim Verde As Integer
Dim Azul As Integer

  Azul = (Color And 16711680) / 65536
  Verde = (Color And 65280) / 256
  Rojo = Color And 255
  
RGBtoD3DColorARGB = D3DColorARGB(Alpha, Rojo, Verde, Azul)
  
End Function

Public Sub Device_Box_Textured_Render(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)

    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long

    
    If GrhIndex = 0 Then Exit Sub
    Set d3dTextures.Texture = GetTexture(GrhIndex, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
    
    If (light_value(0) = 0) Then light_value(0) = base_light
    If (light_value(1) = 0) Then light_value(1) = base_light
    If (light_value(2) = 0) Then light_value(2) = base_light
    If (light_value(3) = 0) Then light_value(3) = base_light
        
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .left = src_x
        .Right = src_x + src_width
        .top = src_y
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .left = dest_x
        .Right = dest_x + src_width
        .top = dest_y
    End With
    
    
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle


    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.Texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    'D3DDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLESTRIP, 0, 4, 2, indexList(0), D3DFMT_INDEX16, temp_verts(0), Len(temp_verts(0))
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub

Private Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal Color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.Color = Color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function
Public Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Integer, Optional ByRef Textures_Height As Integer, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
   
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.left + (dest.Right - dest.left) / 2
        y_center = dest.top + (dest.bottom - dest.top) / 2
       
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
       
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = 3.1459 - right_point
    End If
   
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
   
   
    '0 - Bottom left vertex
    If Textures_Width Or Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.left / Textures_Width + 0.001, (src.bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
   
   
    '1 - Top left vertex
    If Textures_Width Or Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.left / Textures_Width + 0.001, src.top / Textures_Height + 0.001)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
   
   
    '2 - Bottom right vertex
    If Textures_Width Or Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
   
   
    '3 - Top right vertex
    If Textures_Width Or Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.top / Textures_Height + 0.001)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)
    End If
 
End Sub
Public Function GetTexture(ByVal FileName As Integer, ByRef textwidth As Integer, ByRef textheight As Integer) As Direct3DTexture8
'WWWW.RINCONDELAO.COM.AR
If FileName = 0 Then
Debug.Print "0 GRH ATMPT TO BE LOADED"
Exit Function
End If
 
    Dim i As Long
    ' Search the index on the list
    With TexList(FileName Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).FileName = FileName Then
                .SurfaceEntry(i).UltimoAcceso = GetTickCount
                textwidth = .SurfaceEntry(i).texture_width
                textheight = .SurfaceEntry(i).texture_height
                Set GetTexture = .SurfaceEntry(i).Texture
                Exit Function
            End If
        Next i
    End With
 
    'Not in memory, load it!
    Set GetTexture = CrearGrafico(FileName, textwidth, textheight)
End Function
Private Function CrearGrafico(ByVal Archivo As Integer, ByRef texwidth As Integer, ByRef textheight As Integer) As Direct3DTexture8
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'menduz was here
'
'**************************************************************
On Error GoTo ErrHandler
    Dim surface_desc As D3DSURFACE_DESC
    Dim texture_info As D3DXIMAGE_INFO
    Dim Index As Integer
    Index = Archivo Mod HASH_TABLE_SIZE
    With TexList(Index)
        .surfaceCount = .surfaceCount + 1
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        With .SurfaceEntry(.surfaceCount)
            'Nombre
            .FileName = Archivo
           
            'Ultimo acceso
            .UltimoAcceso = GetTickCount
   
            Set .Texture = mD3D.CreateTextureFromFileEx(device, App.Path & "\GRAFICOS\" & LTrim(Str(Archivo)) & ".bmp", _
                D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, &HFF000000, texture_info, ByVal 0)
               
            .Texture.GetLevelDesc 0, surface_desc
            .texture_width = texture_info.Width
            .texture_height = texture_info.Height
            .size = surface_desc.size
            texwidth = .texture_width
            textheight = .texture_height
            Set CrearGrafico = .Texture
            mFreeMemoryBytes = mFreeMemoryBytes + surface_desc.size
        End With
    End With
    Debug.Print mFreeMemoryBytes / 1024 / 1024; " MB LIBRES"
    Do While mFreeMemoryBytes < 0
        If Not RemoveLRU() Then
            Exit Do
        End If
    Loop
Exit Function
ErrHandler:
Debug.Print "ERROR EN GRHLOAD>" & Archivo & ".bmp"
End Function

Private Function RemoveLRU() As Boolean
'**************************************************************
'Author: Juan Mart?n Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Removes the Least Recently Used surface to make some room for new ones
'WWWW.RINCONDELAO.COM.AR
'**************************************************************
    Dim LRUi As Long
    Dim LRUj As Long
    Dim LRUtime As Long
    Dim i As Long
    Dim j As Long
    Dim surface_desc As D3DSURFACE_DESC
   
    LRUtime = GetTickCount
   
    'Check out through the whole list for the least recently used
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                If LRUtime > .SurfaceEntry(j).UltimoAcceso Then
                    LRUi = i
                    LRUj = j
                    LRUtime = .SurfaceEntry(j).UltimoAcceso
                End If
            Next j
        End With
    Next i
   
    'Retrieve the surface desc
    Call TexList(LRUi).SurfaceEntry(LRUj).Texture.GetLevelDesc(0, surface_desc)
   
    'Remove it
    Set TexList(LRUi).SurfaceEntry(LRUj).Texture = Nothing
    TexList(LRUi).SurfaceEntry(LRUj).FileName = 0
   
    'Move back the list (if necessary)
    If LRUj Then
        RemoveLRU = True
       
        With TexList(LRUi)
            For j = LRUj To .surfaceCount - 1
                .SurfaceEntry(j) = .SurfaceEntry(j + 1)
            Next j
           
            .surfaceCount = .surfaceCount - 1
            If .surfaceCount Then
                ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
            Else
                Erase .SurfaceEntry
            End If
        End With
    End If
   
    'Update the used bytes
    mFreeMemoryBytes = mFreeMemoryBytes + surface_desc.size
End Function

Public Sub DrawText(ByVal font As Integer, ByVal left As Long, ByVal top As Long, ByVal Text As String, ByVal Color As Long, Optional ByVal Alpha As Byte = 255, Optional ByVal center As Boolean = False)
'*********************************************************
'****** Coded by Dunkan (emanuel.m@dunkancorp.com) *******
'*********************************************************
    If Alpha <> 255 Then
        Dim aux As D3DCOLORVALUE
        ARGBtoD3DCOLORVALUE Color, aux
        Color = D3DColorARGB(Alpha, aux.r, aux.G, aux.b)
    End If
    If Not blur Then
        Engine_Render_Text cfonts(font), Text, left, top, Color, center, Alpha
    End If
End Sub
Public Function ARGBtoD3DCOLORVALUE(ByVal ARGB As Long, ByRef Color As D3DCOLORVALUE)
Dim dest(3) As Byte
CopyMemory dest(0), ARGB, 4
Color.a = dest(3)
Color.r = dest(2)
Color.G = dest(1)
Color.b = dest(0)
End Function
 
 
Private Sub Engine_Render_Text(ByRef UseFont As CustomFont, ByVal Text As String, ByVal X As Long, ByVal Y As Long, ByVal Color As Long, Optional ByVal center As Boolean = False, Optional ByVal Alpha As Byte = 255)
Dim TempVA(0 To 3) As TLVERTEX
Dim tempstr() As String
Dim count As Integer
Dim ascii() As Byte
Dim Row As Integer
Dim u As Single
Dim v As Single
Dim i As Long
Dim j As Long
Dim KeyPhrase As Byte
Dim TempColor As Long
Dim ResetColor As Byte
Dim SrcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim YOffset As Single
Dim bucleFonts As Integer

    For bucleFonts = 1 To UBound(cfonts)
 
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    'D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
   
    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
 
    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
   
    'Get the text into arrays (split by vbCrLf)
    tempstr = Split(Text, vbCrLf)
   
    'Set the temp color (or else the first character has no color)
    TempColor = Color
 
    'Set the texture
    D3DDevice.SetTexture 0, UseFont.Texture
   
    If center Then
        X = X - Engine_GetTextWidth(cfonts(bucleFonts), Text) * 0.5
    End If
   
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(tempstr)
        If Len(tempstr(i)) > 0 Then
            YOffset = i * UseFont.CharHeight
            count = 0
       
            'Convert the characters to the ascii value
            ascii() = StrConv(tempstr(i), vbFromUnicode)
       
            'Loop through the characters
            For j = 1 To Len(tempstr(i))
 
                'Check for a key phrase
                'If ascii(j - 1) = 124 Then 'If Ascii = "|"
                '    KeyPhrase = (Not KeyPhrase)  'TempColor = ARGB 255/255/0/0
                '    If KeyPhrase Then TempColor = ARGB(255, 0, 0, alpha) Else ResetColor = 1
                'Else
 
                    'Render with triangles
                    'If AlternateRender = 0 Then
 
                        'Copy from the cached vertex array to the temp vertex array
                        CopyMemory TempVA(0), UseFont.HeaderInfo.CharVA(ascii(j - 1)).Vertex(0), 32 * 4
 
                        'Set up the verticies
                        TempVA(0).X = X + count
                        TempVA(0).Y = Y + YOffset
                       
                        TempVA(1).X = TempVA(1).X + X + count
                        TempVA(1).Y = TempVA(0).Y
 
                        TempVA(2).X = TempVA(0).X
                        TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
 
                        TempVA(3).X = TempVA(1).X
                        TempVA(3).Y = TempVA(2).Y
                       
                        'Set the colors
                        TempVA(0).Color = TempColor
                        TempVA(1).Color = TempColor
                        TempVA(2).Color = TempColor
                        TempVA(3).Color = TempColor
                       
                        'Draw the verticies
                        'D3DDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLESTRIP, 0, 4, 2, indexList(0), D3DFMT_INDEX16, TempVA(0), Len(TempVA(0))
                        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TempVA(0), Len(TempVA(0))
                     
                    'Shift over the the position to render the next character
                    count = count + UseFont.HeaderInfo.CharWidth(ascii(j - 1))
               
                'End If
               
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = Color
                End If
               
            Next j
           
        End If
    Next i
    
Next bucleFonts
   
End Sub

Private Function Engine_GetTextWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Integer
'***************************************************
'Returns the width of text
'More info: http://www.vbgore.com/GameClient.TileEn ... tTextWidth
'***************************************************
Dim i As Integer
 
    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
   
    'Loop through the text
    For i = 1 To Len(Text)
       
        'Add up the stored character widths
        Engine_GetTextWidth = Engine_GetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(mid$(Text, i, 1)))
       
    Next i
 
End Function
 
Sub Engine_Init_FontTextures()
On Error GoTo eDebug:
'*****************************************************************
'Init the custom font textures
'More info: http://www.vbgore.com/GameClient.TileEn ... ntTextures
'*****************************************************************
Dim TexInfo As D3DXIMAGE_INFO_A
 
    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
 
    '*** Default font ***
   
    'Set the texture
    Dim bucleFonts As Integer
    For bucleFonts = 1 To UBound(cfonts)
    Set cfonts(bucleFonts).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Init\Fuentes\" & bucleFonts & ".bmp", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)
   
    'Store the size of the texture
    cfonts(bucleFonts).TextureSize.X = TexInfo.Width
    cfonts(bucleFonts).TextureSize.Y = TexInfo.Height
    Next bucleFonts
   
    Exit Sub
eDebug:
    If Err.Number = "-2005529767" Then
        MsgBox "Error en la carga de las fuentes.", vbCritical
        End
    End If
    End
 
End Sub
 
Sub Engine_Init_FontSettings()
'*********************************************************
'****** Coded by Dunkan (emanuel.m@dunkancorp.com) *******
'*********************************************************
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single
Dim bucleFonts As Integer
 
For bucleFonts = 1 To UBound(cfonts)
 
    'Load the header information
    FileNum = FreeFile
    Open App.Path & "\Init\Fuentes\" & bucleFonts & ".dat" For Binary As #FileNum
        Get #FileNum, , cfonts(bucleFonts).HeaderInfo
    Close #FileNum
   
    'Calculate some common values
    cfonts(bucleFonts).CharHeight = cfonts(bucleFonts).HeaderInfo.CellHeight - 4
    cfonts(bucleFonts).RowPitch = cfonts(bucleFonts).HeaderInfo.BitmapWidth \ cfonts(bucleFonts).HeaderInfo.CellWidth
    cfonts(bucleFonts).ColFactor = cfonts(bucleFonts).HeaderInfo.CellWidth / cfonts(bucleFonts).HeaderInfo.BitmapWidth
    cfonts(bucleFonts).RowFactor = cfonts(bucleFonts).HeaderInfo.CellHeight / cfonts(bucleFonts).HeaderInfo.BitmapHeight
   
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
       
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - cfonts(bucleFonts).HeaderInfo.BaseCharOffset) \ cfonts(bucleFonts).RowPitch
        u = ((LoopChar - cfonts(bucleFonts).HeaderInfo.BaseCharOffset) - (Row * cfonts(bucleFonts).RowPitch)) * cfonts(bucleFonts).ColFactor
        v = Row * cfonts(bucleFonts).RowFactor
 
        'Set the verticies
        With cfonts(bucleFonts).HeaderInfo.CharVA(LoopChar)
            .Vertex(0).Color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).rhw = 1
            .Vertex(0).tu = u
            .Vertex(0).tv = v
            .Vertex(0).X = 0
            .Vertex(0).Y = 0
            .Vertex(0).Z = 0
           
            .Vertex(1).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).rhw = 1
            .Vertex(1).tu = u + cfonts(bucleFonts).ColFactor
            .Vertex(1).tv = v
            .Vertex(1).X = cfonts(bucleFonts).HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).Z = 0
           
            .Vertex(2).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).rhw = 1
            .Vertex(2).tu = u
            .Vertex(2).tv = v + cfonts(bucleFonts).RowFactor
            .Vertex(2).X = 0
            .Vertex(2).Y = cfonts(bucleFonts).HeaderInfo.CellHeight
            .Vertex(2).Z = 0
           
            .Vertex(3).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).rhw = 1
            .Vertex(3).tu = u + cfonts(bucleFonts).ColFactor
            .Vertex(3).tv = v + cfonts(bucleFonts).RowFactor
            .Vertex(3).X = cfonts(bucleFonts).HeaderInfo.CellWidth
            .Vertex(3).Y = cfonts(bucleFonts).HeaderInfo.CellHeight
            .Vertex(3).Z = 0
        End With
       
    Next LoopChar
    Next bucleFonts

End Sub

Sub DDrawTono(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, _
    Optional ByVal KillAnim As Integer, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, _
    Optional reflejoagua As Boolean, Optional Alpha As Boolean, Optional RadioBrillo As Byte = 0, Optional ColorBrillo As Long, Optional angle As Single, Optional Color As Long)
 
    Dim iGrhIndex As Integer
    Dim QuitarAnimacion As Boolean
 
    If Grh.GrhIndex = 0 Then Exit Sub
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.SpeedCounter)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
               
                If Grh.SpeedCounter <> -1 Then
                    If Grh.SpeedCounter > 0 Then
                        Grh.SpeedCounter = Grh.SpeedCounter - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    If iGrhIndex = 0 Then Exit Sub
    
    If center Then
        If GrhData(iGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16
        End If
        
        If GrhData(iGrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32
        End If
    End If

    Dim C(3) As Long
     
    ColorBrillo = D3DColorXRGB(255, 255, 255)
 
    C(0) = ColorBrillo
    C(1) = ColorBrillo
    C(2) = ColorBrillo
    C(3) = ColorBrillo
        
    Device_Box_Textured_Render_Advance iGrhIndex, _
        X + RadioBrillo, Y, _
        GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
        C(), _
        GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
        True, angle, reflejoagua

    Device_Box_Textured_Render_Advance iGrhIndex, _
        X - RadioBrillo, Y, _
        GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
        C(), _
        GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
        True, angle, reflejoagua
    
    Device_Box_Textured_Render_Advance iGrhIndex, _
        X, Y + RadioBrillo, _
        GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
        C(), _
        GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
        True, angle, reflejoagua
    
    Device_Box_Textured_Render_Advance iGrhIndex, _
        X, Y - RadioBrillo, _
        GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
        C(), _
        GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
        True, angle, reflejoagua
    
    Device_Box_Textured_Render_Advance iGrhIndex, _
        X, Y, _
        GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
        C(), _
        GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
        Alpha, angle, reflejoagua
    
    
 
End Sub

Private Sub Device_Box_Textured_Render_Advance(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single, _
                                            Optional ByVal shadow As Byte = 0)
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long
   
    If GrhIndex = 0 Then Exit Sub
    Set d3dTextures.Texture = GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
   
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
   
    If (light_value(0) = 0) Then light_value(0) = base_light
    If (light_value(1) = 0) Then light_value(1) = base_light
    If (light_value(2) = 0) Then light_value(2) = base_light
    If (light_value(3) = 0) Then light_value(3) = base_light
       
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .left = src_x
        .Right = src_x + src_width
        .top = src_y
    End With
               
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .left = dest_x
        .Right = dest_x + src_width
        .top = dest_y
    End With
   
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
   
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.Texture
   
    If shadow Then
        temp_verts(1).X = temp_verts(1).X + src_width / 2
        temp_verts(1).Y = temp_verts(1).Y - src_height / 2
     
        temp_verts(3).X = temp_verts(3).X + src_width
        temp_verts(3).Y = temp_verts(3).Y - src_width
    End If
   
    If alpha_blend Then
      'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
   
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
   
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
 
End Sub

