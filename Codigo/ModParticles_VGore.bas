Attribute VB_Name = "modParticles"
'/////////////////////////////Motor Grafico en DirectX 8///////////////////////////////
'////////////////////////Extraccion de varios motores por ShaFTeR//////////////////////
'///////////////////ORE - VBGORE - GSZAO - KKAO y algunos ejemplos de webs/////////////
'**************************************************************************************

Option Explicit
'Texture for particle effects - this is handled differently then the rest of the graphics
Public ParticleTexture(1 To 12) As Direct3DTexture8
Public ParticleIndex() As Integer

Private Type Effect
    X As Single                 'Location of effect
    Y As Single
    GoToX As Single             'Location to move to
    GoToY As Single
    KillWhenAtTarget As Boolean     'If the effect is at its target (GoToX/Y), then Progression is set to 0
    KillWhenTargetLost As Boolean   'Kill the effect if the target is lost (sets progression = 0)
    Gfx As Byte                 'Particle texture used
    Used As Boolean             'If the effect is in use
    EffectNum As Byte           'What number of effect that is used
    Modifier As Integer         'Misc variable (depends on the effect)
    FloatSize As Long           'The size of the particles
    Direction As Integer        'Misc variable (depends on the effect)
    Particles() As Particle     'Information on each particle
    Progression As Single       'Progression state, best to design where 0 = effect ends
    PartVertex() As TLVERTEX    'Used to point render particles
    PreviousFrame As Long       'Tick time of the last frame
    ParticleCount As Integer    'Number of particles total
    ParticlesLeft As Integer    'Number of particles left - only for non-repetitive effects
    BindToChar As Integer       'Setting this value will bind the effect to move towards the character
    BindSpeed As Single         'How fast the effect moves towards the character
    BoundToMap As Byte          'If the effect is bound to the map or not (used only by the map editor)
    r As Single
    G As Single
    b As Single
    EcuationCount As Byte
End Type

Public NumEffects As Byte   'Maximum number of effects at once
Public Effect() As Effect   'List of all the active effects

'Constants With The Order Number For Each Effect
Public Const EffectNum_Fire As Byte = 1             'Burn baby, burn! Flame from a central point that blows in a specified direction
Public Const EffectNum_Snow As Byte = 2             'Snow that covers the screen - weather effect
Public Const EffectNum_Heal As Byte = 3             'Healing effect that can bind to a character, ankhs float up and fade
Public Const EffectNum_Bless As Byte = 4            'Following three effects are same: create a circle around the central point
Public Const EffectNum_Protection As Byte = 5       ' (often the character) and makes the given particle on the perimeter
Public Const EffectNum_Strengthen As Byte = 6       ' which float up and fade out
Public Const EffectNum_Rain As Byte = 7             'Exact same as snow, but moves much faster and more alpha value - weather effect
Public Const EffectNum_EquationTemplate As Byte = 8 'Template for creating particle effects through equations - a page with some equations can be found here: http://www.vbgore.com/modules.php?name=Forums&file=viewtopic&t=221
Public Const EffectNum_Waterfall As Byte = 9        'Waterfall effect
Public Const EffectNum_Summon As Byte = 10          'Summon effect
Public Const EffectNum_Meditate As Byte = 11        'Medit effect
Public Const EffectNum_Portal As Byte = 12          'Portal effect
Public Const EffectNum_Atomic As Byte = 13          'Atomic Effect
Public Const EffectNum_Circle As Byte = 14          'Outlined Circle Effect
Public Const EffectNum_Raro As Byte = 15
Public Const EffectNum_Lissajous As Byte = 16
Public Const EffectNum_Apocalipsis As Byte = 17
Public Const EffectNum_Humo As Byte = 18
Public Const EffectNum_CherryBlossom As Byte = 19
Public Const EffectNum_BloodSpray As Byte = 20
Public Const EffectNum_BloodSplatter As Byte = 21
Public Const EffectNum_LevelUP As Byte = 22         'Level Up Effect
Public Const EffectNum_AnimatedSign As Byte = 23
Public Const EffectNum_Galaxy As Byte = 24
Public Const EffectNum_FancyThickCircle As Byte = 25
Public Const EffectNum_Flower As Byte = 26
Public Const EffectNum_Wormhole As Byte = 27
Public Const EffectNum_HouseTeleport As Byte = 28   'Teleport To House Effect
Public Const EffectNum_GuildTeleport As Byte = 29   'Teleport To Guild Meeting
Public Const EffectNum_Rayo As Byte = 30             'Tormenta de Fuego
Public Const EffectNum_LissajousMedit As Byte = 31

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal length As Long)

Sub Engine_Init_ParticleEngine(Optional ByVal SkipToTextures As Boolean = False)
'*****************************************************************
'Loads all particles into memory - unlike normal textures, these stay in memory. This isn't
'done for any reason in particular, they just use so little memory since they are so small
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_ParticleEngine
'*****************************************************************
Dim i As Byte

    If Not SkipToTextures Then
        'Set the particles texture
        NumEffects = 30 'General_Var_Get(App.Path & "\Game.ini", "INIT", "NumEffects")
        ReDim Effect(1 To NumEffects)
    End If
    
    For i = 1 To UBound(ParticleTexture())
        If Not ParticleTexture(i) Is Nothing Then Set ParticleTexture(i) = Nothing
        Set ParticleTexture(i) = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Graficos\Particles\" & i & ".png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)
    Next i

ReDim ParticleIndex(1) As Integer
    
End Sub



Function Effect_EquationTemplate_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'Particle effect template for effects as described on the
'wiki page: http://www.vbgore.com/Particle_effect_equations
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_EquationTemplate_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_EquationTemplate  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_EquationTemplate_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_EquationTemplate_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Reset
'*****************************************************************
Dim X As Single
Dim Y As Single
Dim r As Single
    
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.1
    r = (Index / 20) * Exp(Index / Effect(EffectIndex).Progression Mod 3)
    X = r * Cos(Index)
    Y = r * Sin(Index)
    
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 1, 0.2 + (Rnd * 0.2)

End Sub

Private Sub Effect_EquationTemplate_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go Through The Particle Loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_EquationTemplate_Reset EffectIndex, loopc

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub

Function Effect_Bless_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Bless_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Bless     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Bless_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Bless_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Reset
'*****************************************************************
Dim a As Single
Dim X As Single
Dim Y As Single

    'Get the positions
    a = Rnd * 360 * DegreeToRadian
    X = Effect(EffectIndex).X - (Sin(a) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Cos(a) * Effect(EffectIndex).Modifier)

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)

End Sub

Private Sub Effect_Bless_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go Through The Particle Loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Bless_Reset EffectIndex, loopc

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub

Function Effect_Fire_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Fire_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Fire      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Fire_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Fire_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Reset
'*****************************************************************

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, Cos((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 1, 0.2, 0.2, 0.4 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)

End Sub

Private Sub Effect_Fire_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
    
    'Go Through The Particle Loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Fire_Reset EffectIndex, loopc

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub

Private Function Effect_FToDW(f As Single) As Long
'*****************************************************************
'Converts a float to a D-Word, or in Visual Basic terms, a Single to a Long
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_FToDW
'*****************************************************************
Dim buf As D3DXBuffer

    'Converts a single into a long (Float to DWORD)
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, Effect_FToDW

End Function

Function Effect_Heal_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Heal_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Heal      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Progression = Progression   'Loop the effect
    Effect(EffectIndex).KillWhenAtTarget = True     'End the effect when it reaches the target (progression = 0)
    Effect(EffectIndex).KillWhenTargetLost = True   'End the effect if the target is lost (progression = 0)
    
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Heal_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Heal_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Reset
'*****************************************************************

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), Cos((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 0.8, 0.2, 0.2, 0.6 + (Rnd * 0.2), 0.01 + (Rnd * 0.5)
    
End Sub

Private Sub Effect_Heal_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression <> 0 Then

                    'Reset the particle
                    Effect_Heal_Reset EffectIndex, loopc

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0

                End If

            Else
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub

Sub Effect_Kill(ByVal EffectIndex As Integer, Optional ByVal KillAll As Boolean = False)
'*****************************************************************
'Kills (stops) a single effect or all effects
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Kill
'*****************************************************************
Dim loopc As Long

    'Check If To Kill All Effects
    If KillAll = True Then

        'Loop Through Every Effect
        For loopc = 1 To NumEffects

            'Stop The Effect
            Effect(loopc).Used = False

        Next
        
    Else

        'Stop The Selected Effect
        Effect(EffectIndex).Used = False
        
    End If

End Sub

Private Function Effect_NextOpenSlot() As Integer
'*****************************************************************
'Finds the next open effects index
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_NextOpenSlot
'*****************************************************************
Dim EffectIndex As Integer

    'Find The Next Open Effect Slot
    Do
        EffectIndex = EffectIndex + 1   'Check The Next Slot
        If EffectIndex > NumEffects Then    'Dont Go Over Maximum Amount
            Effect_NextOpenSlot = -1
            Exit Function
        End If
    Loop While Effect(EffectIndex).Used = True    'Check Next If Effect Is In Use

    'Return the next open slot
    Effect_NextOpenSlot = EffectIndex

    'Clear the old information from the effect
    Erase Effect(EffectIndex).Particles()
    Erase Effect(EffectIndex).PartVertex()
    ZeroMemory Effect(EffectIndex), LenB(Effect(EffectIndex))
    Effect(EffectIndex).GoToX = -30000
    Effect(EffectIndex).GoToY = -30000

End Function

Function Effect_Protection_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Protection_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Protection    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Protection_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Protection_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Reset
'*****************************************************************
Dim a As Single
Dim X As Single
Dim Y As Single

    'Get the positions
    a = Rnd * 360 * DegreeToRadian
    X = Effect(EffectIndex).X - (Sin(a) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Cos(a) * Effect(EffectIndex).Modifier)

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.1, 0.9, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)

End Sub

Private Sub Effect_UpdateOffset(ByVal EffectIndex As Integer)
'***************************************************
'Update an effect's position if the screen has moved
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateOffset
'***************************************************

    Effect(EffectIndex).X = Effect(EffectIndex).X + (LastOffsetX - ParticleOffsetX)
    Effect(EffectIndex).Y = Effect(EffectIndex).Y + (LastOffsetY - ParticleOffsetY)

End Sub

Private Sub Effect_UpdateBinding(ByVal EffectIndex As Integer)

'***************************************************
'Updates the binding of a particle effect to a target, if
'the effect is bound to a character
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateBinding
'***************************************************
Dim TargetI As Integer
Dim TargetA As Single
 
    'Update position through character binding
    If Effect(EffectIndex).BindToChar > 0 Then
 
        'Store the character index
        TargetI = Effect(EffectIndex).BindToChar
 
        'Check for a valid binding index
        If TargetI > LastChar Then
            Effect(EffectIndex).BindToChar = 0
            If Effect(EffectIndex).KillWhenTargetLost Then
                Effect(EffectIndex).Progression = 0
                Exit Sub
            End If
        ElseIf CharList(TargetI).active = 0 Then
            Effect(EffectIndex).BindToChar = 0
            If Effect(EffectIndex).KillWhenTargetLost Then
                Effect(EffectIndex).Progression = 0
                Exit Sub
            End If
        Else
 
            'Calculate the X and Y positions
            Effect(EffectIndex).GoToX = Engine_TPtoSPX(CharList(Effect(EffectIndex).BindToChar).POS.X)
            Effect(EffectIndex).GoToY = Engine_TPtoSPY(CharList(Effect(EffectIndex).BindToChar).POS.Y)
 
        End If
 
    End If
 
    'Move to the new position if needed
    If Effect(EffectIndex).GoToX > -30000 Or Effect(EffectIndex).GoToY > -30000 Then
        If Effect(EffectIndex).GoToX <> Effect(EffectIndex).X Or Effect(EffectIndex).GoToY <> Effect(EffectIndex).Y Then
 
            'Calculate the angle
            TargetA = Engine_GetAngle(Effect(EffectIndex).X, Effect(EffectIndex).Y, Effect(EffectIndex).GoToX, Effect(EffectIndex).GoToY) + 180
 
            'Update the position of the effect
            Effect(EffectIndex).X = Effect(EffectIndex).X - Sin(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed * timerElapsedTime
            Effect(EffectIndex).Y = Effect(EffectIndex).Y + Cos(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed * timerElapsedTime
 
            'Check if the effect is close enough to the target to just stick it at the target
            If Effect(EffectIndex).GoToX > -30000 Then
                If Abs(Effect(EffectIndex).X - Effect(EffectIndex).GoToX) < 2 Then Effect(EffectIndex).X = Effect(EffectIndex).GoToX
            End If
            If Effect(EffectIndex).GoToY > -30000 Then
                If Abs(Effect(EffectIndex).Y - Effect(EffectIndex).GoToY) < 2 Then Effect(EffectIndex).Y = Effect(EffectIndex).GoToY
            End If
 
            'Check if the position of the effect is equal to that of the target
            If Effect(EffectIndex).X = Effect(EffectIndex).GoToX Then
                If Effect(EffectIndex).Y = Effect(EffectIndex).GoToY Then
 
                    'For some effects, if the position is reached, we want to end the effect
                    If Effect(EffectIndex).KillWhenAtTarget Then
                        Effect(EffectIndex).BindToChar = 0
                        Effect(EffectIndex).Progression = 0
                        Effect(EffectIndex).GoToX = Effect(EffectIndex).X
                        Effect(EffectIndex).GoToY = Effect(EffectIndex).Y
                    End If
                    Exit Sub    'The effect is at the right position, don't update
 
                End If
            End If
 
        End If
    End If
 
End Sub


Private Sub Effect_Protection_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Protection_Reset EffectIndex, loopc

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub

Public Sub Effect_Render(ByVal EffectIndex As Integer, Optional ByVal SetRenderStates As Boolean = True)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Render
'*****************************************************************
Dim count As Long
Dim i As Long

    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Set the render state for the size of the particle
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Effect(EffectIndex).FloatSize
    
    'Set the render state to point blitting
    If SetRenderStates Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    
    'Set the last texture to a random number to force the engine to reload the texture
    'LastTexture = -65489

    'Check what type of rendering to do (blood or everything else)
    If Effect(EffectIndex).EffectNum = EffectNum_BloodSpray Or Effect(EffectIndex).EffectNum = EffectNum_BloodSplatter Then

        count = Effect(EffectIndex).ParticleCount \ 4

        D3DDevice.SetTexture 0, ParticleTexture(12)
        'D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, Count, Effect(EffectIndex).PartVertex(0), Len(Effect(EffectIndex).PartVertex(0))
        
        D3DDevice.DrawIndexedPrimitiveUP D3DPT_POINTLIST, 0, 4, count, _
                            indexList(0), D3DFMT_INDEX16, _
                            Effect(EffectIndex).PartVertex(0), Len(Effect(EffectIndex).PartVertex(0))

        For i = 0 To count - 1
            With Effect(EffectIndex).Particles(i)
                If .sngZ < 1 Then Effect(EffectIndex).PartVertex(i).Y = Effect(EffectIndex).PartVertex(i).Y + .sngZ
                Effect(EffectIndex).PartVertex(i).Color = D3DColorMake(.SngR, .SngG, .SngB, .SngA)
            End With
        Next i
        'D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, Count, Effect(EffectIndex).PartVertex(0), Len(Effect(EffectIndex).PartVertex(0))

        D3DDevice.DrawIndexedPrimitiveUP D3DPT_POINTLIST, 0, 4, count, _
                            indexList(0), D3DFMT_INDEX16, _
                            Effect(EffectIndex).PartVertex(0), Len(Effect(EffectIndex).PartVertex(0))

        D3DDevice.SetTexture 0, ParticleTexture(12)
        'D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, Count, Effect(EffectIndex).PartVertex(Count - 1), Len(Effect(EffectIndex).PartVertex(0))
        
        D3DDevice.DrawIndexedPrimitiveUP D3DPT_POINTLIST, 0, 4, count, _
                            indexList(0), D3DFMT_INDEX16, _
                            Effect(EffectIndex).PartVertex(count - 1), Len(Effect(EffectIndex).PartVertex(0))

        For i = count To count - 1 + count
            With Effect(EffectIndex).Particles(i)
                If .sngZ < 1 Then Effect(EffectIndex).PartVertex(i).Y = Effect(EffectIndex).PartVertex(i).Y + .sngZ
                Effect(EffectIndex).PartVertex(i).Color = D3DColorMake(.SngR, .SngG, .SngB, .SngA)
            End With
        Next i
        'D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, Count, Effect(EffectIndex).PartVertex(Count - 1), Len(Effect(EffectIndex).PartVertex(0))

        D3DDevice.DrawIndexedPrimitiveUP D3DPT_POINTLIST, 0, 4, count, _
                            indexList(0), D3DFMT_INDEX16, _
                            Effect(EffectIndex).PartVertex(count - 1), Len(Effect(EffectIndex).PartVertex(0))


        D3DDevice.SetTexture 0, ParticleTexture(12)
        'D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, Count, Effect(EffectIndex).PartVertex((Count * 2) - 1), Len(Effect(EffectIndex).PartVertex(0))
        
        D3DDevice.DrawIndexedPrimitiveUP D3DPT_POINTLIST, 0, 4, count, _
                            indexList(0), D3DFMT_INDEX16, _
                            Effect(EffectIndex).PartVertex((count * 2) - 1), Len(Effect(EffectIndex).PartVertex(0))

        For i = (count * 2) To (count * 2) - 1 + count
            With Effect(EffectIndex).Particles(i)
                If .sngZ < 1 Then Effect(EffectIndex).PartVertex(i).Y = Effect(EffectIndex).PartVertex(i).Y + .sngZ
                Effect(EffectIndex).PartVertex(i).Color = D3DColorMake(.SngR, .SngG, .SngB, .SngA)
            End With
        Next i
        
        'D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, Count, Effect(EffectIndex).PartVertex((Count * 2) - 1), Len(Effect(EffectIndex).PartVertex(0))
        D3DDevice.DrawIndexedPrimitiveUP D3DPT_POINTLIST, 0, 4, count, _
                            indexList(0), D3DFMT_INDEX16, _
                            Effect(EffectIndex).PartVertex((count * 2) - 1), Len(Effect(EffectIndex).PartVertex(0))

        D3DDevice.SetTexture 0, ParticleTexture(12)
        'D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, Count, Effect(EffectIndex).PartVertex((Count * 3) - 1), Len(Effect(EffectIndex).PartVertex(0))
        
        D3DDevice.DrawIndexedPrimitiveUP D3DPT_POINTLIST, 0, 4, count, _
                            indexList(0), D3DFMT_INDEX16, _
                            Effect(EffectIndex).PartVertex((count * 3) - 1), Len(Effect(EffectIndex).PartVertex(0))

        For i = (count * 3) To Effect(EffectIndex).ParticleCount
            With Effect(EffectIndex).Particles(i)
                If .sngZ < 1 Then Effect(EffectIndex).PartVertex(i).Y = Effect(EffectIndex).PartVertex(i).Y + .sngZ
                Effect(EffectIndex).PartVertex(i).Color = D3DColorMake(.SngR, .SngG, .SngB, .SngA)
            End With
        Next i
        
        'D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, Count, Effect(EffectIndex).PartVertex((Count * 3) - 1), Len(Effect(EffectIndex).PartVertex(0))
        
        D3DDevice.DrawIndexedPrimitiveUP D3DPT_POINTLIST, 0, 4, count, _
                            indexList(0), D3DFMT_INDEX16, _
                            Effect(EffectIndex).PartVertex((count * 3) - 1), Len(Effect(EffectIndex).PartVertex(0))

    Else

    'Set the texture
    D3DDevice.SetTexture 0, ParticleTexture(Effect(EffectIndex).Gfx)

    'Draw all the particles at once
    D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, Effect(EffectIndex).ParticleCount, Effect(EffectIndex).PartVertex(0), Len(Effect(EffectIndex).PartVertex(0))

    'Reset the render state back to normal
    If SetRenderStates Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    End If

End Sub

Function Effect_Snow_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Snow_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Snow      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Snow_Reset EffectIndex, loopc, 1
    Next loopc

    'Set the initial time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Snow_Reset(ByVal EffectIndex As Integer, ByVal Index As Long, Optional ByVal FirstReset As Byte = 0)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Reset
'*****************************************************************

    If FirstReset = 1 Then

        'The very first reset
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (frmMain.ScaleWidth + 400)), Rnd * (frmMain.ScaleHeight + 50), Rnd * 5, 5 + Rnd * 3, 0, 0

    Else

        'Any reset after first
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (frmMain.ScaleWidth + 400)), -15 - Rnd * 185, Rnd * 5, 5 + Rnd * 3, 0, 0
        If Effect(EffectIndex).Particles(Index).sngX < -20 Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (frmMain.ScaleHeight + 50)
        If Effect(EffectIndex).Particles(Index).sngX > frmMain.ScaleWidth Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (frmMain.ScaleHeight + 50)
        If Effect(EffectIndex).Particles(Index).sngY > frmMain.ScaleHeight Then Effect(EffectIndex).Particles(Index).sngX = Rnd * (frmMain.ScaleWidth + 50)

    End If

    'Set the color
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 0.8, 0

End Sub

Private Sub Effect_Snow_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check if particle is in use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if to reset the particle
            If Effect(EffectIndex).Particles(loopc).sngX < -200 Then Effect(EffectIndex).Particles(loopc).SngA = 0
            If Effect(EffectIndex).Particles(loopc).sngX > (frmMain.ScaleWidth + 200) Then Effect(EffectIndex).Particles(loopc).SngA = 0
            If Effect(EffectIndex).Particles(loopc).sngY > (frmMain.ScaleHeight + 200) Then Effect(EffectIndex).Particles(loopc).SngA = 0

            'Time for a reset, baby!
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Reset the particle
                Effect_Snow_Reset EffectIndex, loopc

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub

Function Effect_Strengthen_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Strengthen_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Strengthen    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Strengthen_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Strengthen_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Reset
'*****************************************************************
Dim a As Single
Dim X As Single
Dim Y As Single

    'Get the positions
    a = Rnd * 360 * DegreeToRadian
    X = Effect(EffectIndex).X - (Sin(a) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Cos(a) * Effect(EffectIndex).Modifier)

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 0.2, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)

End Sub

Private Sub Effect_Strengthen_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check if particle is in use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update the particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Strengthen_Reset EffectIndex, loopc

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub

Sub Effect_UpdateAll()
'*****************************************************************
'Updates all of the effects and renders them
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateAll
'*****************************************************************
Dim loopc As Long

    'Make sure we have effects
    If NumEffects = 0 Then Exit Sub

    'Set the render state for the particle effects
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    'Update every effect in use
    For loopc = 1 To NumEffects

        'Make sure the effect is in use
        If Effect(loopc).Used Then
        
            'Update the effect position if the screen has moved
            Effect_UpdateOffset loopc
        
            'Update the effect position if it is binded
            Effect_UpdateBinding loopc

            'Find out which effect is selected, then update it
            If Effect(loopc).EffectNum = EffectNum_Fire Then Effect_Fire_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Snow Then Effect_Snow_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Heal Then Effect_Heal_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Bless Then Effect_Bless_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Protection Then Effect_Protection_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Strengthen Then Effect_Strengthen_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Rain Then Effect_Rain_Update loopc
            If Effect(loopc).EffectNum = EffectNum_EquationTemplate Then Effect_EquationTemplate_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Waterfall Then Effect_Waterfall_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Summon Then Effect_Summon_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Meditate Then Effect_Meditate_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Portal Then Effect_Portal_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Atomic Then Effect_Atomic_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Circle Then Effect_Circle_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Raro Then Effect_Raro_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Lissajous Then Effect_Lissajous_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Apocalipsis Then Effect_Apocalipsis_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Humo Then Effect_Humo_Update loopc
            If Effect(loopc).EffectNum = EffectNum_CherryBlossom Then Effect_CherryBlossom_Update loopc
            If Effect(loopc).EffectNum = EffectNum_BloodSpray Then Effect_BloodSpray_Update loopc
            If Effect(loopc).EffectNum = EffectNum_BloodSplatter Then Effect_BloodSplatter_Update loopc
            If Effect(loopc).EffectNum = EffectNum_LevelUP Then Effect_Spawn_Update EffectNum_LevelUP, loopc
            If Effect(loopc).EffectNum = EffectNum_AnimatedSign Then Effect_Spawn_Update EffectNum_AnimatedSign, loopc
            If Effect(loopc).EffectNum = EffectNum_Galaxy Then Effect_Spawn_Update EffectNum_Galaxy, loopc
            If Effect(loopc).EffectNum = EffectNum_FancyThickCircle Then Effect_Spawn_Update EffectNum_FancyThickCircle, loopc
            If Effect(loopc).EffectNum = EffectNum_Flower Then Effect_Spawn_Update EffectNum_Flower, loopc
            If Effect(loopc).EffectNum = EffectNum_Wormhole Then Effect_Spawn_Update EffectNum_Wormhole, loopc
            If Effect(loopc).EffectNum = EffectNum_HouseTeleport Then Effect_Spawn_Update EffectNum_HouseTeleport, loopc
            If Effect(loopc).EffectNum = EffectNum_GuildTeleport Then Effect_Spawn_Update EffectNum_GuildTeleport, loopc
            If Effect(loopc).EffectNum = EffectNum_Rayo Then Effect_Rayo_Update loopc
            If Effect(loopc).EffectNum = EffectNum_LissajousMedit Then Effect_LissajousMedit_Update loopc
            'Render the effect
            Effect_Render loopc, False

        End If

    Next
    
    'Set the render state back for normal rendering
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

End Sub

Function Effect_Rain_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Rain_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Rain      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(10)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Rain_Reset EffectIndex, loopc, 1
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Rain_Reset(ByVal EffectIndex As Integer, ByVal Index As Long, Optional ByVal FirstReset As Byte = 0)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Reset
'*****************************************************************

    If FirstReset = 1 Then

        'The very first reset
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (frmMain.ScaleWidth + 400)), Rnd * (frmMain.ScaleHeight + 50), Rnd * 5, 25 + Rnd * 12, 0, 0

    Else

        'Any reset after first
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * 1200), -15 - Rnd * 185, Rnd * 5, 25 + Rnd * 12, 0, 0
        If Effect(EffectIndex).Particles(Index).sngX < -20 Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (frmMain.ScaleHeight + 50)
        If Effect(EffectIndex).Particles(Index).sngX > frmMain.ScaleWidth Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (frmMain.ScaleHeight + 50)
        If Effect(EffectIndex).Particles(Index).sngY > frmMain.ScaleHeight Then Effect(EffectIndex).Particles(Index).sngX = Rnd * (frmMain.ScaleWidth + 50)

    End If

    'Set the color
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 0.4, 0

End Sub

Private Sub Effect_Rain_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check if the particle is in use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update the particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if to reset the particle
            If Effect(EffectIndex).Particles(loopc).sngX < -200 Then Effect(EffectIndex).Particles(loopc).SngA = 0
            If Effect(EffectIndex).Particles(loopc).sngX > (frmMain.ScaleWidth + 200) Then Effect(EffectIndex).Particles(loopc).SngA = 0
            If Effect(EffectIndex).Particles(loopc).sngY > (frmMain.ScaleHeight + 200) Then Effect(EffectIndex).Particles(loopc).SngA = 0

            'Time for a reset, baby!
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Reset the particle
                Effect_Rain_Reset EffectIndex, loopc

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub

Public Sub Effect_Begin(ByVal EffectIndex As Integer, ByVal X As Single, ByVal Y As Single, ByVal GfxIndex As Byte, ByVal Particles As Byte, Optional ByVal Direction As Single = 180, Optional ByVal BindToMap As Boolean = False)
'*****************************************************************
'A very simplistic form of initialization for particle effects
'Should only be used for starting map-based effects
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Begin
'*****************************************************************
Dim RetNum As Byte

    Select Case EffectIndex
        Case EffectNum_Fire
            RetNum = Effect_Fire_Begin(X, Y, GfxIndex, Particles, Direction, 1)
        Case EffectNum_Waterfall
            RetNum = Effect_Waterfall_Begin(X, Y, GfxIndex, Particles)
    End Select
    
    'Bind the effect to the map if needed
    If BindToMap Then Effect(RetNum).BoundToMap = 1
    
End Sub

Function Effect_Waterfall_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Waterfall_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Waterfall     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Waterfall_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Waterfall_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Reset
'*****************************************************************

    If Int(Rnd * 10) = 1 Then
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + (Rnd * 60), Effect(EffectIndex).Y + (Rnd * 130), 0, 8 + (Rnd * 6), 0, 0
    Else
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + (Rnd * 60), Effect(EffectIndex).Y + (Rnd * 10), 0, 8 + (Rnd * 6), 0, 0
    End If
    Effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.1, 0.9, 0.6 + (Rnd * 0.4), 0
    
End Sub

Private Sub Effect_Waterfall_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount
    
        With Effect(EffectIndex).Particles(loopc)
    
            'Check if the particle is in use
            If .Used Then
    
                'Update The Particle
                .UpdateParticle ElapsedTime

                'Check if the particle is ready to die
                If (.sngY > Effect(EffectIndex).Y + 140) Or (.SngA = 0) Then
    
                    'Reset the particle
                    Effect_Waterfall_Reset EffectIndex, loopc
    
                Else

                    'Set the particle information on the particle vertex
                    Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(.SngR, .SngG, .SngB, .SngA)
                    Effect(EffectIndex).PartVertex(loopc).X = .sngX
                    Effect(EffectIndex).PartVertex(loopc).Y = .sngY
    
                End If
    
            End If
            
        End With

    Next loopc

End Sub

Function Effect_Summon_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 0) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Summon_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Summon    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Summon_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Summon_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Reset
'*****************************************************************
Dim X As Single
Dim Y As Single
Dim r As Single
    
    If Effect(EffectIndex).Progression > 1000 Then
        Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 1.4
    Else
        Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.5
    End If
    r = (Index / 30) * Exp(Index / Effect(EffectIndex).Progression)
    X = r * Cos(Index)
    Y = r * Sin(Index)
    
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 0, Rnd, 0, 0.9, 0.2 + (Rnd * 0.2)
 
End Sub

Private Sub Effect_Summon_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go Through The Particle Loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression < 1800 Then

                    'Reset the particle
                    Effect_Summon_Reset EffectIndex, loopc

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0

                End If

            Else
            
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub

Function Effect_Meditate_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... tate_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Meditate_Begin = EffectIndex
 
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Meditate     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Meditate_Reset EffectIndex, loopc
    Next loopc
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Meditate_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... tate_Reset
'*****************************************************************
Dim a As Single
Dim X As Single
Dim Y As Single
Dim rR As Single
Dim RG As Single
Dim rB As Single

   'Get the positions
   a = Rnd * 360 * DegreeToRadian
   X = Effect(EffectIndex).X - (Sin(a) * Effect(EffectIndex).Modifier)
   Y = Effect(EffectIndex).Y + (Cos(a) * Effect(EffectIndex).Modifier / 2.5)
   
   'Load Colours
   rR = (0.1 - 0.05) * Rnd + 0.03
   RG = 0.8
   rB = 0.5
 
   'Reset the particle
   Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
   Effect(EffectIndex).Particles(Index).ResetColor rR, RG, rB, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
   
End Sub
 
Private Sub Effect_Meditate_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... ate_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long
 
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
 
    'Go Through The Particle Loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount
 
        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then
 
            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime
 
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then
 
                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then
 
                    'Reset the particle
                    Effect_Meditate_Reset EffectIndex, loopc
 
                Else
 
                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False
 
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
 
                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False
 
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0
 
                End If
 
            Else
 
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY
 
            End If
 
        End If
 
    Next loopc
 
End Sub

Function Effect_Portal_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... rtal_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Portal_Begin = EffectIndex
 
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Portal     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Portal_Reset EffectIndex, loopc
    Next loopc
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Portal_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... rtal_Reset
'*****************************************************************
Dim a As Single
Dim X As Single
Dim Y As Single
Dim rR As Single
Dim RG As Single
Dim rB As Single
 
   'Get the positions
   a = Rnd * 360 * DegreeToRadian
   If Rnd > Rnd Then
    X = Effect(EffectIndex).X - (Sin(a) * Effect(EffectIndex).Modifier / 1.8) * Rnd * 1.1
    Y = Effect(EffectIndex).Y + (Cos(a) * Effect(EffectIndex).Modifier * 1.1) * Rnd * 1.1
    rR = (0.1 - 0.05) * Rnd + 0.03
    RG = 0.2
    rB = 0.8
   Else
    X = Effect(EffectIndex).X - (Sin(a) * Effect(EffectIndex).Modifier / 3)
    Y = Effect(EffectIndex).Y + (Cos(a) * Effect(EffectIndex).Modifier / 1.5)
    rR = (0.2 - 0.06) * Rnd + 0.04
    RG = 0.3
    rB = 0.2
   End If
 
   'Reset the particle
   Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, 0 '-2
   Effect(EffectIndex).Particles(Index).ResetColor rR, RG, rB, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
 
End Sub
 
Private Sub Effect_Portal_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... tal_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long
 
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
 
    'Go Through The Particle Loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount
 
        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then
 
            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime
 
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then
 
                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then
 
                    'Reset the particle
                    Effect_Portal_Reset EffectIndex, loopc
 
                Else
 
                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False
 
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
 
                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False
 
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0
 
                End If
 
            Else
 
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY
 
            End If
 
        End If
 
    Next loopc
 
End Sub

Function Effect_Atomic_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************

'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Atomic_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Atomic        'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Atomic_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Atomic_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************

'*****************************************************************
Dim r As Single
Dim X As Single
Dim Y As Single

    'Get the positions
    r = 10 + Sin(2 * (Index / 10)) * 50
    X = r * Cos(Index / 30)
    Y = r * Sin(Index / 30)

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 200, 50, 1, 1, 0.9 + (Rnd * 0.2)
End Sub

Private Sub Effect_Atomic_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Atomic_Reset EffectIndex, loopc

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub


Function Effect_Circle_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************

'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Circle_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Circle        'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Circle_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Circle_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************

'*****************************************************************
Dim a As Single
Dim X As Single
Dim Y As Single

    'Get the positions
    a = Rnd * 360 * DegreeToRadian 'The point on the circumference to be used
    X = Effect(EffectIndex).X - (Sin(a) * 40) 'The 40s state the radius of circle
    Y = Effect(EffectIndex).Y + (Cos(a) * 40)

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 1 * Rnd + 0.4, 0, 1, 1, 0.2 + (Rnd * 0.2)

End Sub

Private Sub Effect_Circle_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Circle_Reset EffectIndex, loopc

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub

Function Effect_Raro_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************

'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Raro_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Raro        'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Raro_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Raro_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************

'*****************************************************************
Dim X As Single
Dim Y As Single
Dim i As Single
    'Get the positions
    'a = Rnd * 360 * DegreeToRadian 'The point on the circumference to be used
    For i = 0 To 360 Step 30
    X = Effect(EffectIndex).X - Cos(i)
    Y = Effect(EffectIndex).Y + Sin(i) + Rnd

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 1, 0.2 + (Rnd * 0.2)
    
    Next i
End Sub

Private Sub Effect_Raro_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Raro_Reset EffectIndex, loopc

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub

Function Effect_Lissajous_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 0, Optional size As Byte = 30, Optional r As Single = 100, Optional G As Single = 100, Optional b As Single = 100)
'*****************************************************************
'Particle effect Lissajous equation
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Lissajous_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Lissajous 'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Modifier = size                 'How large the circle is
    Effect(EffectIndex).Progression = Progression
    Effect(EffectIndex).r = r
    Effect(EffectIndex).G = G
    Effect(EffectIndex).b = b
    
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Lissajous_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Lissajous_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... late_Reset
'*****************************************************************
Dim X As Single
Dim Y As Single
Dim a As Single
Dim e1 As Byte
Dim e2 As Byte
Dim e3 As Byte
Dim e4 As Byte
Dim s1 As Byte 'suma1 de la segunda ecuacion
Dim s2 As Byte
 
e1 = 2
e2 = 1
e3 = 1
e4 = 2
s1 = 5
s2 = 7
 
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.01
   
    a = Effect(EffectIndex).Progression
   
    If RandomNumber(1, 2) = 1 Then
        X = Effect(EffectIndex).X - (Sin(e1 * a) * Effect(EffectIndex).Modifier) - 20
        Y = Effect(EffectIndex).Y + (Sin(e2 * a) * Effect(EffectIndex).Modifier)
        'Reset the particle
        Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, 0
        Effect(EffectIndex).Particles(Index).ResetColor Effect(EffectIndex).r * Effect(EffectIndex).Progression, Effect(EffectIndex).G * Effect(EffectIndex).Progression, Effect(EffectIndex).b, 0.2, 0.2 + (Rnd * 0.2)
 
    Else
        X = Effect(EffectIndex).X - (Sin(e3 + s1 * a) * Effect(EffectIndex).Modifier) - 20
        Y = Effect(EffectIndex).Y + (Sin(e4 + s2 * a) * Effect(EffectIndex).Modifier)
        'Reset the particle
        Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, 0
        Effect(EffectIndex).Particles(Index).ResetColor Effect(EffectIndex).r * Effect(EffectIndex).Progression, Effect(EffectIndex).G * Effect(EffectIndex).Progression, Effect(EffectIndex).b, 0.2, 0.2 + (Rnd * 0.2)
 
    End If
   
End Sub
 

Private Sub Effect_Lissajous_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go Through The Particle Loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Lissajous_Reset EffectIndex, loopc

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub

Function Effect_Apocalipsis_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 0) As Integer
'*****************************************************************
'Particle effect template for effects as described on the
'wiki page: http://www.vbgore.com/Particle_effect_equations
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Apocalipsis_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Apocalipsis  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression
    'Effect(EffectIndex).KillWhenAtTarget = True
    
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Apocalipsis_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Apocalipsis_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Reset
'*****************************************************************
Dim X As Single
Dim Y As Single
Dim a As Single

    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.01
    
    a = Effect(EffectIndex).Progression
    'If RandomNumber(1, 2) = 1 Then
    X = Effect(EffectIndex).X '- (Sin(a)) * 120
    Y = Effect(EffectIndex).Y '+ Cos(5 * a) * 20 'The 40s state the radius of circle
    
    Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 5, 0, 3, 1, 0.2 + (Rnd * 0.2)
    'Else
    'x = Effect(EffectIndex).x - (Sin(a)) * 120
    'y = Effect(EffectIndex).y - Cos(5 * a) * 20 'The 40s state the radius of circle
    '
    'Effect(EffectIndex).Particles(Index).ResetIt x, y, 0, 0, 0, 0
    'Effect(EffectIndex).Particles(Index).ResetColor 0, 5, 2, 1, 0.2 + (Rnd * 0.2)
    'End If
    
End Sub

Private Sub Effect_Apocalipsis_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go Through The Particle Loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Apocalipsis_Reset EffectIndex, loopc

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub

Function Effect_Humo_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://svn2.assembla.com/svn/vblore/trunk/Code/Common%20Code/Particles.bas
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Humo_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Humo      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(30)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Humo_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Humo_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://svn2.assembla.com/svn/vblore/trunk/Code/Common%20Code/Particles.bas
'*****************************************************************

    'Reset the particle
    'Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, Cos((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, 0, 0
    'Effect(EffectIndex).Particles(Index).ResetColor 1, 0.2, 0.2, 0.4 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 50, Effect(EffectIndex).Y - 10 + Rnd * 50, -Sin((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 5, Cos((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, 0.5, 0
    Effect(EffectIndex).Particles(Index).ResetColor 0.2, 0.2, 0.2, 0.2 + (Rnd * 0.2), 0.03 + (Rnd * 0.01)

    'Reset the particle
    'Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 50, Effect(EffectIndex).Y - 10 + Rnd * 80, -Sin((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, Cos((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, 0, 0
    'Effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.1, 0.1, 0.4 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)
    'Effect(EffectIndex).Particles(index).ResetColor 0.1, 0.1, 0.1, 0.4 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)

End Sub

Private Sub Effect_Humo_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://svn2.assembla.com/svn/vblore/trunk/Code/Common%20Code/Particles.bas
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go Through The Particle Loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression <> 0 Then

                    'Reset the particle
                    Effect_Humo_Reset EffectIndex, loopc

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub

Function Effect_CherryBlossom_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_CherryBlossom_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_CherryBlossom     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_CherryBlossom_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_CherryBlossom_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Reset
'*****************************************************************

    If Int(Rnd * 10) = 1 Then
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + (Rnd * 60), Effect(EffectIndex).Y + (Rnd * 130), 2 + (Rnd * 2), 2 + (Rnd * 2), 0, 0
    Else
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + (Rnd * 60), Effect(EffectIndex).Y + (Rnd * 10), 2 + (Rnd * 2), 2 + (Rnd * 2), 0, 0
    End If
    Effect(EffectIndex).Particles(Index).ResetColor 1#, 0.7, 0.75, 0.6 + (Rnd * 0.4), 0
    
End Sub

Private Sub Effect_CherryBlossom_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount
    
        With Effect(EffectIndex).Particles(loopc)
    
            'Check if the particle is in use
            If .Used Then
    
                'Update The Particle
                .UpdateParticle ElapsedTime

                'Check if the particle is ready to die
                If (.sngY > Effect(EffectIndex).Y + 140) Or (.SngA = 0) Then
    
                    'Reset the particle
                    Effect_CherryBlossom_Reset EffectIndex, loopc
    
                Else

                    'Set the particle information on the particle vertex
                    Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(.SngR, .SngG, .SngB, .SngA)
                    Effect(EffectIndex).PartVertex(loopc).X = .sngX
                    Effect(EffectIndex).PartVertex(loopc).Y = .sngY
    
                End If
    
            End If
            
        End With

    Next loopc

End Sub

Function Effect_BloodSpray_Begin(ByVal X As Single, ByVal Y As Single, ByVal Particles As Integer, ByVal Direction As Single, Optional ByVal Intensity As Single = 1) As Integer
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_BloodSpray_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_BloodSpray  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Direction = Direction           'Direction
    Effect(EffectIndex).Modifier = Intensity

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(7)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_BloodSpray_Reset EffectIndex, loopc
    Next loopc

    'Set the initial time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_BloodSpray_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

    'Reset the particle
    With Effect(EffectIndex)
        .Particles(Index).ResetIt .X + (Rnd * 16) - 8, .Y + (Rnd * 32) - 16, _
             Sin((.Direction - 10 + (Rnd * 20)) * DegreeToRadian) * (30 * .Modifier * Rnd), _
            -Cos((.Direction - 10 + (Rnd * 20)) * DegreeToRadian) * (30 * .Modifier * Rnd), 0, 0, -10, -2 - (Rnd * 30), 8 + Rnd * 4
        .Particles(Index).ResetColor 1, 1, 1, 0.8, 0
    End With
    
End Sub

Private Sub Effect_BloodSpray_Update(ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim loopc As Long
Dim TileX As Long
Dim TileY As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount
    
        With Effect(EffectIndex).Particles(loopc)
    
            'Check if particle is in Use
            If .Used Then
    
                'Update the particle
                .UpdateParticle ElapsedTime
                
                'Don't pass any walls/etc
                TileX = Engine_SPtoTPX(.sngX)
                TileY = Engine_SPtoTPY(.sngY)
                If TileX < 1 Then
                    .sngZ = 1.1
                ElseIf TileY < 1 Then
                    .sngZ = 1.1
                ElseIf TileX > 92 Then
                    .sngZ = 1.1
                ElseIf TileY > 92 Then
                    .sngZ = 1.1
                End If
                If .sngZ <> 1.1 Then
                    'If MapData(TileX, TileY).BlockedAttack Then
                        '.sngZ = 1.1
                    'End If
                End If
                
                'Blood trails
                If loopc = 0 Or loopc Mod 15 = 0 Then
                    If Int(Rnd * 3) = 0 Then
                        If Int(Rnd * 2) = 0 Then
                            Engine_Blood_Create .sngX + ParticleOffsetX, .sngY + ParticleOffsetY
                        Else
                            Engine_Blood_Create .sngX + ParticleOffsetX, .sngY + ParticleOffsetY
                        End If
                    End If
                End If
    
                'Check if to kill off the particle
                If .sngZ > 1 Then

                    'Disable the particle
                    .Used = False
    
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
    
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).Particles(loopc).SngA = 0
                    
                    'Check if we lost all the particles
                    If Effect(EffectIndex).ParticlesLeft <= 0 Then Effect(EffectIndex).Used = False
                    
                    'Create the blood splatter
                    Engine_Blood_Create .sngX + ParticleOffsetX, .sngY + ParticleOffsetY
    
                Else

                    'Set the particle information on the particle vertex
                    Effect(EffectIndex).PartVertex(loopc).Color = 1258291200
                    Effect(EffectIndex).PartVertex(loopc).X = .sngX
                    Effect(EffectIndex).PartVertex(loopc).Y = .sngY
    
                End If
    
            End If
            
        End With

    Next loopc

End Sub

Function Effect_BloodSplatter_Begin(ByVal X As Single, ByVal Y As Single, ByVal Particles As Integer) As Integer
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_BloodSplatter_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_BloodSplatter  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(7)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_BloodSplatter_Reset EffectIndex, loopc
    Next loopc

    'Set the initial time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_BloodSplatter_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
Dim Direction As Single

    'Find the direction
    Direction = Rnd * 360

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + (Rnd * 16) - 8, Effect(EffectIndex).Y + (Rnd * 32) - 16, _
         Sin(Direction * DegreeToRadian) * (24 * Rnd), _
        -Cos(Direction * DegreeToRadian) * (24 * Rnd), 0, 0, -25, -3 - (Rnd * 40), 10 + Rnd * 4
    Effect(EffectIndex).Particles(Index).ResetColor 1, 0, 0, 0.8, 0
End Sub

Private Sub Effect_BloodSplatter_Update(ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim loopc As Long
Dim TileX As Long
Dim TileY As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount
    
        With Effect(EffectIndex).Particles(loopc)
    
            'Check if particle is in Use
            If .Used Then
    
                'Update the particle
                .UpdateParticle ElapsedTime
                
                'Don't pass any walls/etc
                TileX = Engine_SPtoTPX(.sngX)
                TileY = Engine_SPtoTPY(.sngY)
                If TileX < 1 Then
                    .sngZ = 1.1
                ElseIf TileY < 1 Then
                    .sngZ = 1.1
                ElseIf TileY > 92 Then
                    .sngZ = 1.1
                ElseIf TileY > 92 Then
                    .sngZ = 1.1
                End If
                If .sngZ <> 1.1 Then
                    'If MapData(TileX, TileY).BlockedAttack Then
                        '.sngZ = 1.1
                    'End If
                End If
                
                'Blood trails
                If loopc = 0 Or loopc Mod 10 = 0 Then
                    If Int(Rnd * 3) = 0 Then
                        If Int(Rnd * 2) = 0 Then
                            Engine_Blood_Create .sngX + ParticleOffsetX, .sngY + ParticleOffsetY
                        Else
                            Engine_Blood_Create .sngX + ParticleOffsetX, .sngY + ParticleOffsetY
                        End If
                    End If
                End If
    
                'Check if to kill off the particle
                If .sngZ > 1 Then
                
                    'Disable the particle
                    .Used = False
    
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
    
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).Particles(loopc).SngA = 0
                    
                    'Check if we lost all the particles
                    If Effect(EffectIndex).ParticlesLeft <= 0 Then Effect(EffectIndex).Used = False
                    
                    'Create the blood splatter
                    Engine_Blood_Create .sngX + ParticleOffsetX, .sngY + ParticleOffsetY
    
                Else

                    'Set the particle information on the particle vertex
                    Effect(EffectIndex).PartVertex(loopc).Color = 1258291200
                    Effect(EffectIndex).PartVertex(loopc).X = .sngX
                    Effect(EffectIndex).PartVertex(loopc).Y = .sngY
    
                End If
    
            End If
            
        End With

    Next loopc

End Sub


Function Effect_Spawn_Begin(ByVal EffectNum As Byte, ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal size As Byte = 30, Optional ByVal Time As Single = 10, Optional ByVal Red As Single = -1, Optional ByVal Green As Single = -1, Optional ByVal Blue As Single = -1, Optional ByVal Alpha As Single = -1) As Integer
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Spawn_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Spawn_Reset EffectNum, EffectIndex, loopc, Red, Green, Blue, Alpha
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Spawn_Reset(ByVal EffectNum As Byte, ByVal EffectIndex As Integer, ByVal Index As Long, Optional ByVal Red As Single = -1, Optional ByVal Green As Single = -1, Optional ByVal Blue As Single = -1, Optional ByVal Alpha As Single = -1)
Dim a As Single
Dim b As Single
Dim X As Single
Dim Y As Single
Dim r As Single


    'Determine if deafults are used
    If Red = -2 Then Red = Rnd
    If Green = -2 Then Green = Rnd
    If Blue = -2 Then Blue = Rnd
    If Alpha = -2 Then Alpha = Rnd
    
    
    'store
    Effect(EffectIndex).Particles(Index).Red = Red
    Effect(EffectIndex).Particles(Index).Green = Green
    Effect(EffectIndex).Particles(Index).Blue = Blue
    Effect(EffectIndex).Particles(Index).Alpha = Alpha
    
    Select Case EffectNum
        Case EffectNum_HouseTeleport
            r = Sin(20 / (Index + 1)) * 100
            X = r * Cos((Index))
            Y = r * Sin((Index))
            
            'Reset the particle
            Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
            
            'Determine if deafults are used
            If Red = -1 Then Red = Rnd
            If Green = -1 Then Green = Rnd
            If Blue = -1 Then Blue = 1
            If Alpha = -1 Then Alpha = Rnd
            
            Effect(EffectIndex).Particles(Index).ResetColor Red, Green, Blue, Alpha, 0.2 + (Rnd * 0.5)
        Case EffectNum_GuildTeleport
            r = 150 + Cos(Index * Rnd) * Sin(Index * Rnd)
            X = r * Cos(Index) * Rnd
            Y = r * Sin(Index) * Rnd
            
            
            'Determine if deafults are used
            If Red = -1 Then Red = Rnd
            If Green = -1 Then Green = Rnd
            If Blue = -1 Then Blue = 0.5
            If Alpha = -1 Then Alpha = Rnd
            
            'Reset the particle
            Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
            Effect(EffectIndex).Particles(Index).ResetColor Red, Green, Blue, Alpha, 0.2 + (Rnd * 0.2)
        Case EffectNum_LevelUP
            r = 10 + Sin(2 * (Index / 10)) * 50 + (30 * Rnd)
            X = r * Cos(Index / 30)
            Y = r * Sin(Index / 30)
            
            'Determine if deafults are used
            If Red = -1 Then Red = 1
            If Green = -1 Then Green = 0.3 + Rnd / 2
            If Blue = -1 Then Blue = Rnd / 3
            If Alpha = -1 Then Alpha = Rnd / 2
        
           'Reset the particle
           Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
           Effect(EffectIndex).Particles(Index).ResetColor Red, Green, Blue, Alpha, 0.005 + (Rnd * 0.2)
        Case EffectNum_AnimatedSign
            If Index = 0 Then Effect(EffectIndex).Modifier = Effect(EffectIndex).Modifier + 1
            Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + Effect(EffectIndex).Direction
            If Effect(EffectIndex).Progression > 100 Then Effect(EffectIndex).Direction = -0.02
            If Effect(EffectIndex).Progression < -100 Then Effect(EffectIndex).Direction = 0.02
         
            r = Effect(EffectIndex).Progression + 2 * Cos(2 * Index) * 40
            X = r * Cos(Index / (Effect(EffectIndex).Modifier + 1) * 5)
            Y = r * Sin(Index / (Effect(EffectIndex).Modifier + 1) * 5)
        
            'Determine if deafults are used
            If Red = -1 Then Red = 1
            If Green = -1 Then Green = 1
            If Blue = -1 Then Blue = 1
            If Alpha = -1 Then Alpha = 1
            
            'Reset the particle
            Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
            Effect(EffectIndex).Particles(Index).ResetColor Red, Green, Blue, Alpha, 0.2 + (Rnd * 0.2)
        Case EffectNum_Galaxy
            r = Sin(20 / (Index + 1)) * 100
            X = r * Cos((Index))
            Y = r * Sin((Index))
        
            'Determine if deafults are used
            If Red = -1 Then Red = 1
            If Green = -1 Then Green = 1
            If Blue = -1 Then Blue = 1
            If Alpha = -1 Then Alpha = 1
            
            'Reset the particle
            Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
            Effect(EffectIndex).Particles(Index).ResetColor Red, Green, Blue, Alpha, 0.2 + (Rnd * 0.2)
        Case EffectNum_FancyThickCircle
           
            r = 50 + Rnd * 15 * Cos(2 * Index)
            X = r * Cos(Index / 30)
            Y = r * Sin(Index / 30)
        
            'Determine if deafults are used
            If Red = -1 Then Red = 1
            If Green = -1 Then Green = 1
            If Blue = -1 Then Blue = 1
            If Alpha = -1 Then Alpha = 1
            
            'Reset the particle
            Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
            Effect(EffectIndex).Particles(Index).ResetColor Red, Green, Blue, Alpha, 0.2 + (Rnd * 0.2)
        Case EffectNum_Flower
            r = Cos(2 * (Index / 10)) * 50
            X = r * Cos(Index / 10)
            Y = r * Sin(Index / 10)
        
            'Determine if deafults are used
            If Red = -1 Then Red = 1
            If Green = -1 Then Green = 1
            If Blue = -1 Then Blue = 1
            If Alpha = -1 Then Alpha = 1
            
            'Reset the particle
            Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
            Effect(EffectIndex).Particles(Index).ResetColor Red, Green, Blue, Alpha, 0.2 + (Rnd * 0.2)
        Case EffectNum_Wormhole
            Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.1
            r = (Index / 20) * Exp(Index / Effect(EffectIndex).Progression Mod 3)
            X = r * Cos(Index)
            Y = r * Sin(Index)
        
            'Determine if deafults are used
            If Red = -1 Then Red = 1
            If Green = -1 Then Green = 1
            If Blue = -1 Then Blue = 1
            If Alpha = -1 Then Alpha = 1
            
            'Reset the particle
            Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
            Effect(EffectIndex).Particles(Index).ResetColor Red, Green, Blue, Alpha, 0.2 + (Rnd * 0.2)
    End Select
End Sub

Private Sub Effect_Spawn_Update(ByVal EffectNum As Byte, ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go Through The Particle Loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Spawn_Reset EffectNum, EffectIndex, loopc, Effect(EffectIndex).Particles(loopc).Red, Effect(EffectIndex).Particles(loopc).Green, Effect(EffectIndex).Particles(loopc).Blue, Effect(EffectIndex).Particles(loopc).Alpha

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY
    
            End If

        End If

    Next loopc

End Sub

Public Sub Effect_Create(ByVal QuienLanza As Byte, ByVal CharIndex As Integer, ByVal Effecto As Byte)

With CharList(CharIndex)

    Select Case Effecto
        Case 1
            .ParticleIndex = Effect_BloodSplatter_Begin(Engine_TPtoSPX(CharList(CharIndex).POS.X), Engine_TPtoSPY(CharList(CharIndex).POS.Y), 20 + Rnd * 40)
        Case 2
            .ParticleIndex = Effect_Rayo_Begin(Engine_TPtoSPX(CharList(QuienLanza).POS.X), Engine_TPtoSPY(CharList(QuienLanza).POS.Y), 13, 100)
            Effect(Effecto).BindToChar = CharIndex
            Effect(Effecto).BindSpeed = 3
    End Select
End With

End Sub

Function Effect_Rayo_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rayo_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Rayo_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Rayo      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Progression = Progression   'Loop the effect
    Effect(EffectIndex).KillWhenAtTarget = True     'End the effect when it reaches the target (progression = 0)
    Effect(EffectIndex).KillWhenTargetLost = True   'End the effect if the target is lost (progression = 0)
    
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Rayo_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Rayo_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rayo_Reset
'*****************************************************************

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), Cos((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 0, 0.8, 0.8, 0.6 + (Rnd * 0.2), 0.001 + (Rnd * 0.5)
'      Effect(EffectIndex).Particles(Index).ResetColor (Rnd * 0.8), (Rnd * 0.8), (Rnd * 0.8), 0.6 + (Rnd * 0.2), 0.001 + (Rnd * 0.5)

End Sub

Private Sub Effect_Rayo_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rayo_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long
Dim i As Integer

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression <> 0 Then

                    'Reset the particle
                    Effect_Rayo_Reset EffectIndex, loopc

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0

                End If

            Else
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub

Function Effect_LissajousMedit_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 0, Optional size As Byte = 30, Optional r As Single = 100, Optional G As Single = 100, Optional b As Single = 100, Optional ByVal EcuationCount = 1)
'*****************************************************************
'Particle effect Lissajous equation
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_LissajousMedit_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_LissajousMedit 'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Modifier = size                 'How large the circle is
    Effect(EffectIndex).Progression = Progression
    Effect(EffectIndex).r = r
    Effect(EffectIndex).G = G
    Effect(EffectIndex).b = b
    Effect(EffectIndex).EcuationCount = EcuationCount
    
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_LissajousMedit_Reset EffectIndex, loopc
    Next loopc

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_LissajousMedit_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Reset
'*****************************************************************
Dim X As Single
Dim Y As Single
Dim a As Single
'2
'1

'1
'2

    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.01
    
    a = Effect(EffectIndex).Progression
With Effect(EffectIndex)

    


If .EcuationCount = 1 Then
        X = Effect(EffectIndex).X - (Sin(1 * a + 1) * Effect(EffectIndex).Modifier) - 20
        Y = Effect(EffectIndex).Y + (Sin(1 * a) * Effect(EffectIndex).Modifier)
        'Reset the particle
        Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, 0
        Effect(EffectIndex).Particles(Index).ResetColor Effect(EffectIndex).r * Effect(EffectIndex).Progression, Effect(EffectIndex).G * Effect(EffectIndex).Progression, Effect(EffectIndex).b, 0.2, 0.2 + (Rnd * 0.2)
ElseIf .EcuationCount = 2 Then
    If RandomNumber(1, 2) = 1 Then
        X = Effect(EffectIndex).X - (Sin(1 * a + 1) * Effect(EffectIndex).Modifier) - 20
        Y = Effect(EffectIndex).Y + (Sin(1 * a) * Effect(EffectIndex).Modifier)
        'Reset the particle
        Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, 0, 0, 0
        Effect(EffectIndex).Particles(Index).ResetColor Effect(EffectIndex).r * Effect(EffectIndex).Progression, Effect(EffectIndex).G * Effect(EffectIndex).Progression, Effect(EffectIndex).b, 0.2, 0.2 + (Rnd * 0.2)

    Else
        X = .X - (Sin(1 * a) * .Modifier) - 20
        Y = .Y + (Sin(1 * a) * .Modifier)
        'Reset the particle
        .Particles(Index).ResetIt X, Y, 0, 0, 0, 0
        .Particles(Index).ResetColor .r * .Progression, .G * .Progression, .b, 0.2, 0.2 + (Rnd * 0.2)

    End If
ElseIf .EcuationCount = 3 Then

    If RandomNumber(1, 2) = 1 Then
        X = .X - (Sin(2 * a) * .Modifier) - 20
        Y = .Y + (Sin(1 * a) * .Modifier)
        'Reset the particle
        .Particles(Index).ResetIt X, Y, 0, 0, 0, 0
        .Particles(Index).ResetColor .r * .Progression, .G * .Progression, .b, 0.2, 0.2 + (Rnd * 0.2)

    Else
        X = .X - (Sin(1 * a) * .Modifier) - 20
        Y = .Y + (Sin(2 * a) * .Modifier)
        'Reset the particle
        .Particles(Index).ResetIt X, Y, 0, 0, 0, 0
        .Particles(Index).ResetColor .r * .Progression, .G * .Progression, .b, 0.2, 0.2 + (Rnd * 0.2)

    End If

ElseIf .EcuationCount = 4 Then

    If RandomNumber(1, 2) = 1 Then
        X = .X - (Sin(4 * a) * .Modifier) - 20
        Y = .Y + (Sin(2 * a) * .Modifier)
        'Reset the particle
        .Particles(Index).ResetIt X, Y, 0, 0, 0, 0
        .Particles(Index).ResetColor .r * .Progression, .G * .Progression, .b, 0.2, 0.2 + (Rnd * 0.2)

    Else
        X = .X - (Sin(2 * a) * .Modifier) - 20
        Y = .Y + (Sin(4 * a) * .Modifier)
        'Reset the particle
        .Particles(Index).ResetIt X, Y, 0, 0, 0, 0
        .Particles(Index).ResetColor .r * .Progression, .G * .Progression, .b, 0.2, 0.2 + (Rnd * 0.2)

    End If

ElseIf .EcuationCount = 5 Then

    If RandomNumber(1, 2) = 1 Then
        X = .X - (Sin(2 * a) * .Modifier) - 20
        Y = .Y + (Sin(1 * a) * .Modifier)
        'Reset the particle
        .Particles(Index).ResetIt X, Y, 0, 0, 0, 0
        .Particles(Index).ResetColor .r * .Progression, .G * .Progression, .b, 0.2, 0.2 + (Rnd * 0.2)

    Else
        X = .X - (Sin(1 + 5 * a) * .Modifier) - 20
        Y = .Y + (Sin(2 + 7 * a) * .Modifier)
        'Reset the particle
        .Particles(Index).ResetIt X, Y, 0, 0, 0, 0
        .Particles(Index).ResetColor .r * .Progression, .G * .Progression, .b, 0.2, 0.2 + (Rnd * 0.2)

    End If
End If

End With
End Sub

Private Sub Effect_LissajousMedit_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go Through The Particle Loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).SngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_LissajousMedit_Reset EffectIndex, loopc

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).Color = D3DColorMake(Effect(EffectIndex).Particles(loopc).SngR, Effect(EffectIndex).Particles(loopc).SngG, Effect(EffectIndex).Particles(loopc).SngB, Effect(EffectIndex).Particles(loopc).SngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY

            End If

        End If

    Next loopc

End Sub

