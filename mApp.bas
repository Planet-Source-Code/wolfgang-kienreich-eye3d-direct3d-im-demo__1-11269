Attribute VB_Name = "mApp"
Option Explicit

 
' APPINITIALIZE: Setup and start
Public Sub AppInitialize()
    
    ' Enable error handling
        On Error GoTo E_AppInitialize
        
    ' Setup local variables ...
    
        Dim L_dDDViewportArea As D3DRECT            ' Area of viewport object for generation of viewport
        Dim L_dDDSurfaceDesc As DDSURFACEDESC2      ' DirectDraw surface description for generation of surfaces
        Dim L_dDDSCAPS As DDSCAPS2                  ' DDCAPS type for backbuffer creation
        Dim L_dM As D3DMATRIX                       ' Matrix for setting up transforms
            
    ' Application specific initialization ...
    
        ' State that initialization is in progress
        G_bAppInitialized = True
        
        ' Initialize statistics
        G_dUser.Stats.Frametime = 0
        
    ' Initialize DirectDraw Instance ...
    
        ' Create instance of DirectDraw
         DirectDrawCreate G_dDXSelectedDriver.GUID, G_oDDInstance, Nothing
        
        ' Check instance existance, terminate if missing
        If G_oDDInstance Is Nothing Then
           AppError 0, "Could not create DirectDraw instance", "AppInitialize"
           Exit Sub
        End If
        
        ' Set DirectDraw cooperative level
        G_oDDInstance.SetCooperativeLevel fApp.hwnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
    
        ' Set display mode
        G_oDDInstance.SetDisplayMode G_nDisplayWidth, G_nDisplayHeight, 16, 0, 0
        
    ' Create primary and backbuffer...
        
        ' For 3DFX, we need a flipping chain ...
        If G_dDXSelectedDriver.DriverType = EDXDTPlus Then
        
            ' Fill surface description
            With L_dDDSurfaceDesc
                .dwSize = Len(L_dDDSurfaceDesc)
                .dwFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
                .DDSCAPS.dwCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_VIDEOMEMORY Or DDSCAPS_3DDEVICE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
                .dwBackBufferCount = 1
            End With
            
            ' Create surface
            G_oDDInstance.CreateSurface L_dDDSurfaceDesc, G_oDDPrimary, Nothing
                                
            ' Check surface existance, terminate if missing
            If G_oDDPrimary Is Nothing Then
               AppError 0, "Could not create primary surface", "AppInitialize"
               Exit Sub
            End If
            
            ' Fill surface description
            With L_dDDSCAPS
                .dwCaps = DDSCAPS_BACKBUFFER
            End With
            
            ' Add surface
            G_oDDPrimary.GetAttachedSurface L_dDDSCAPS, G_oDDBackBuffer

            ' Check surface existance, terminate if missing
            If G_oDDBackBuffer Is Nothing Then
               AppError 0, "Could not create backbuffer surface", "AppInitialize"
               Exit Sub
            End If
            
        ' ... while for standard adapters or accellerators, blitting is enough
        Else
            
            ' Fill surface description
            With L_dDDSurfaceDesc
                .dwSize = Len(L_dDDSurfaceDesc)
                .dwFlags = DDSD_CAPS
                .DDSCAPS.dwCaps = DDSCAPS_PRIMARYSURFACE Or IIf(G_dDXSelectedDriver.DriverType = EDXDTSoft, DDSCAPS_SYSTEMMEMORY, DDSCAPS_VIDEOMEMORY)
            End With
            
            ' Create surface
            G_oDDInstance.CreateSurface L_dDDSurfaceDesc, G_oDDPrimary, Nothing
                                
            ' Check surface existance, terminate if missing
            If G_oDDPrimary Is Nothing Then
               AppError 0, "Could not create primary surface", "AppInitialize"
               Exit Sub
            End If
            
            ' Create surface
            Set G_oDDBackBuffer = CreateSurface(640, 480, DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE)

            ' Check surface existance, terminate if missing
            If G_oDDBackBuffer Is Nothing Then
               AppError 0, "Could not create backbuffer surface", "AppInitialize"
               Exit Sub
            End If
               
        End If
        
    ' Create Direct3DIM environment ...
    
        ' Query DirectDraw for D3D interface
        Set G_oD3DInstance = G_oDDInstance
        
        ' Check instance existance, terminate if missing
        If G_oD3DInstance Is Nothing Then
           AppError 0, "Could not create Direct3D instance", "AppInitialize"
           Exit Sub
        End If
        
        ' Create device using driver found
        G_oD3DInstance.CreateDevice G_dDXSelectedDriver.D3DDriver.GUID, G_oDDBackBuffer, G_oD3DDevice, Nothing
        
        ' Check device existance, terminate if missing
        If G_oD3DDevice Is Nothing Then
           AppError 0, "Could not create Direct3D Device", "AppInitialize"
           Exit Sub
        End If
        
    ' Initialize render viewport ...
    
        ' Create viewport
        G_oD3DInstance.CreateViewport G_oD3DViewport, Nothing
        
        ' Check viewport existance, terminate if missing
        If G_oD3DViewport Is Nothing Then
           AppError 0, "Could not create Direct3D Viewport", "AppInitialize"
           Exit Sub
        End If
        
        ' Add viewport to device
        G_oD3DDevice.AddViewport G_oD3DViewport
        
        ' Set viewport properties
        Call ViewportInitialize(G_nDisplayWidth - G_dUser.DisplaySize, G_nDisplayHeight - G_dUser.DisplaySize)
        
        ' Make viewport current
        G_oD3DDevice.SetCurrentViewport G_oD3DViewport
        
        ' Prepare and set projection transform
        MIdentity L_dM
        L_dM = MProject(1, 130, 60)
        G_oD3DDevice.SetTransform D3DTRANSFORMSTATE_PROJECTION, L_dM
        
        ' Prepare and set world transform
        MIdentity L_dM
        G_oD3DDevice.SetTransform D3DTRANSFORMSTATE_WORLD, L_dM
        
        ' Prepare render and light states
        With G_oD3DDevice
            ' Set perspective correction for textures
            .SetRenderState D3DRENDERSTATE_TEXTUREPERSPECTIVE, IIf(G_dUser.DisplayOptions.Correct, 1, 0)
            ' Set use of specular lighting off
            .SetRenderState D3DRENDERSTATE_SPECULARENABLE, IIf(G_dUser.DisplayOptions.Specular, 1, 0)
            ' Set filtering for textels appearing larger than one pixel
            .SetRenderState D3DRENDERSTATE_TEXTUREMAG, IIf(G_dUser.DisplayOptions.Filtering, D3DFILTER_LINEAR, D3DFILTER_NEAREST)
            ' Set filtering for textels appearing at pixel size or smaller
            .SetRenderState D3DRENDERSTATE_TEXTUREMIN, IIf(G_dUser.DisplayOptions.Filtering, D3DFILTER_LINEAR, D3DFILTER_NEAREST)
            ' Set texure blending options to combine material with texture
            .SetRenderState D3DRENDERSTATE_TEXTUREMAPBLEND, D3DTBLEND_MODULATE
            ' Set alpha blending options
            .SetRenderState D3DRENDERSTATE_SRCBLEND, IIf(G_dUser.DisplayOptions.Translucent, D3DBLEND_BOTHSRCALPHA, D3DBLEND_ONE)
            ' Set texture color key transparency
            .SetRenderState D3DRENDERSTATE_COLORKEYENABLE, IIf(G_dUser.DisplayOptions.Transparent, 1, 0)
            ' Set ambient light level
            .SetLightState D3DLIGHTSTATE_AMBIENT, D3DRMCreateColorRGBA(0.5, 0.5, 0.5, 1)
            ' Set RGB color model ... just to make sure
            .SetLightState D3DLIGHTSTATE_COLORMODEL, D3DCOLOR_RGB
            ' Set shading mode ... just to make sure
            .SetRenderState D3DRENDERSTATE_SHADEMODE, IIf(G_dUser.DisplayOptions.Phong, D3DSHADE_PHONG, D3DSHADE_GOURAUD)
            ' Set fog properties and enable fogging if desired
            .SetLightState D3DLIGHTSTATE_FOGMODE, D3DFOG_NONE
        End With
        
    ' Initialize DirectSound ...
    
        ' Create DirectSound Instance
        DirectSoundCreate ByVal 0&, G_oDSInstance, Nothing
        
        ' Create primary sound buffer
        Set G_oDSBPrimary = CreatePrimaryAudio
        
        ' Play primary buffer
        G_oDSBPrimary.Play ByVal 0&, ByVal 0&, DSBPLAY_LOOPING
        
        ' Get listener
        Set G_oDSListener = G_oDSBPrimary
        
        ' Set listener properties
        G_oDSListener.SetRolloffFactor 1, DS3D_IMMEDIATE
        
    ' Initialize scene data (graphical and audio) ...
        Call SceneInitialize
                
    ' Hide mouse ...
        ShowCursor 0
        
    ' Error handling ...
       
        Exit Sub

E_AppInitialize:
    Resume Next
        AppError Err.Number, Err.Description, "AppInitialize"
        
End Sub

' APPTERMINATE: Cleanup and termination
Public Sub AppTerminate()
    
    ' Enable error handling
        On Error Resume Next
    
    ' Restore from exclusive fullscreen mode ...
        
        ' Restore old resolution and depth
        G_oDDInstance.RestoreDisplayMode
    
        ' Return control to windows
        G_oDDInstance.SetCooperativeLevel fApp.hwnd, DDSCL_NORMAL
        
        ' Show cursor
        ShowCursor 1
            
    ' Clean up graphical data...
        
        Call SceneTerminate
        
    ' Clean up Direct3D...
        
        Set G_oD3DViewport = Nothing
        Set G_oD3DDevice = Nothing
        Set G_oD3DInstance = Nothing
        
    ' Clean up DirectSound...
        G_oDSBPrimary.Stop
        Set G_oDSListener = Nothing
        Set G_oDSBPrimary = Nothing
        Set G_oDSInstance = Nothing
        
    ' Clean up DirectDraw...
    
        Set G_oDDBackBuffer = Nothing
        Set G_oDDPrimary = Nothing
        Set G_oDDInstance = Nothing

    ' Disable error handling
        On Error GoTo 0

End Sub

' APPLOOP: Main program loop
Public Sub AppLoop()

    ' Enable error handling
        On Error GoTo E_AppLoop
        
    ' Setup local variables ...
        Dim L_nRunF As Integer          ' Variable to run through all faces within mesh data
        Dim L_nRunV As Integer          ' Variable to run through vertices within faces
        Dim L_nWaterFactor As Single    ' Texture factor for water texture
        Dim L_dM As D3DMATRIX           ' Matrix to hold various transforms
        Dim L_dV As D3DVECTOR           ' Vector for camera calculation
        Dim L_dDDBLTFX As DDBLTFX       ' FX Blit descriptor
        Dim L_dRenderArea As RECT       ' Blitting area for various blits
        Dim L_dSourceArea As RECT       ' Blitting area for various blits
        Dim L_dDDSD As DDSURFACEDESC2   ' Description of surface to be obtained by lock
        Dim L_nSurfaceDC As Long        ' Pointer to the surface for locking
        
        Dim L_nRunStars As Integer      ' Variable to run through star array
        Dim L_nPosX As Integer          ' Star position after calculations
        Dim L_nPosY As Integer          ' Star position after calculations
        Dim L_nAltitudeFactor As Single ' Altitude factor to correct  star position
        Dim L_nWidthFactor As Single    ' Width Factor to correct star position
       
        Dim L_nCurrentTime As Double    ' Current time for frame timing
        Dim L_nNextSecond As Double     ' Next update time for frame timing monitoring
        Dim L_nNextFrametime As Double  ' Next update time for frame timing monitoring
        Dim L_nFrameCount As Double     ' Frames within update period for frame time monitoring
        
    ' Main application loop ...
        
        ' Set app status to running
        G_bAppRunning = True
        
        Do
    
            ' Do frame timing and statistics ...
            
                ' Increase global frame counter
                G_nFrameCount = G_nFrameCount + 1
            
                ' Get frame start time
                L_nCurrentTime = timeGetTime
                
                ' Increase frame count for avg frametime calculation
                L_nFrameCount = L_nFrameCount + 1
                
                ' Protocol frame time: Count frames and write out average frame count every second
                If L_nNextSecond < L_nCurrentTime Then
                    L_nNextSecond = L_nCurrentTime + 1000
                    G_dUser.Stats.Frametime = L_nFrameCount
                    L_nFrameCount = 0
                End If
            
                ' Prepare timing: Set next frame time to current time plus minimum frame duration (50fps , makes for ~20ms, is max.)
                L_nNextFrametime = L_nCurrentTime + 20
                
            ' D3DIM preparing ...
            
                ' Prepare view transform ...
                    
                    With G_dUser
                    
                        ' Set camera lookat to camera position plus viewing data
                        L_dV.X = .Position.X + Int(Cos(.LookH * PIFactor) * 100)
                        L_dV.z = .Position.z + Int(Sin(.LookH * PIFactor) * 100)
                        L_dV.Y = .Position.Y + 5 + .LookV
                        
                        ' Look there
                        L_dM = MLookAt(.Position, L_dV)
                        G_oD3DDevice.SetTransform D3DTRANSFORMSTATE_VIEW, L_dM
                        
                        ' Listen there
                        G_oDSListener.SetOrientation L_dV.X, L_dV.Y, L_dV.z, 0, -1, 0, DS3D_IMMEDIATE
                        G_oDSListener.SetPosition .Position.X, .Position.Y, .Position.z, DS3D_IMMEDIATE
                        
                    End With
                    
            ' D3DIM rendering ...
                                    
                ' Clear...
                    
                    ' Clear 3D buffer
                    G_oD3DViewport.Clear2 1, G_dClearArea, D3DCLEAR_TARGET, 0, 1, 0
                
'                    ' Clear backbuffer (Necessary for some 3DFX cards !?)
'                    With L_dDDBLTFX
'                        .dwSize = Len(L_dDDBLTFX)
'                        .dwFillColor = 0
'                    End With
'                    G_oDDBackBuffer.Blt G_dClearArea, ByVal Nothing, ByVal 0&, DDBLT_COLORFILL Or DDBLT_WAIT, L_dDDBLTFX
                
                ' Draw background ...
                
                    ' Prepare structure to obtain lock
                    L_dDDSD.dwSize = Len(L_dDDSD)
                    L_dDDSD.dwFlags = DDSD_LPSURFACE

                    ' Obtain lock to surface, get DC
                    G_oDDBackBuffer.GetDC L_nSurfaceDC
                    
                    ' Calculate and draw stars ...
                                            
                         ' Incorporate altitude offset
                         L_nAltitudeFactor = (G_dRenderArea.Bottom - G_dRenderArea.Top) / 107
                         L_nWidthFactor = (G_dRenderArea.Right - G_dRenderArea.Left) / 640
                         
                         ' Run through all stars
                         For L_nRunStars = 0 To 1999
                             
                             ' Evaluate relative position of star
                             With G_dScene.Stars(L_nRunStars)
                                 L_nPosX = .Direction - (G_dUser.LookH * 10) * L_nWidthFactor
                                 If L_nPosX < 0 Then L_nPosX = L_nPosX + 3600
                                 L_nPosY = .Altitude - (G_dUser.LookV * L_nAltitudeFactor)
                             End With
                             
                             ' Draw star if relative position within display area
                             With G_dRenderArea
                                 If L_nPosX > .Left And L_nPosX < .Right And L_nPosY > .Top And L_nPosY < .Bottom Then
                                     SetPixelV L_nSurfaceDC, L_nPosX, L_nPosY, G_dScene.Stars(L_nRunStars).Color
                                 End If
                             End With
                             
                        Next
        
                   ' Release lock to surface
                   G_oDDBackBuffer.ReleaseDC L_nSurfaceDC
                    
                    
                ' Execute polygons onto Direct3DIM...
                With G_oD3DDevice
                    
                    ' Start scene
                    .BeginScene
                    
                    ' Run through vertex data groups
                    For L_nRunF = 0 To UBound(G_dScene.Faces) - 1
                        If G_dScene.Faces(L_nRunF).Enabled Then
                        
                            ' Set group render states (material & transform) ...
                               
                               ' Set transform
                               Select Case L_nRunF
                                
                                    ' Rotating eye (constant rotation)
                                    Case 11
                                        MIdentity L_dM
                                        L_dM = MRotate(L_dM, 0, G_nFrameCount Mod 360, 0)
                                        L_dM = MTranslate(L_dM, 115, 44, 35)
                                    
                                    ' Flame (Decal: Always faces user position)
                                    Case 13
                                        MIdentity L_dM
                                        L_dM = MRotate(L_dM, 0, 180 + Int(Atn((145 - G_dUser.Position.z) / (35 - G_dUser.Position.X)) / PIFactor), 0)
                                        L_dM = MTranslate(L_dM, 47.5, 37, 147.5)
                                        
                                    ' Statics
                                    Case Else
                                        MIdentity L_dM
                                        
                                End Select
                                .SetTransform D3DTRANSFORMSTATE_WORLD, L_dM
                                
                                ' Set material to use
                                    .SetLightState D3DLIGHTSTATE_MATERIAL, G_dScene.Materials(G_dScene.Faces(L_nRunF).D3DMaterialIndex).D3DHandle
                                                      
                                ' Set texture to use
                                    If G_dUser.DisplayOptions.Mapping And Not (G_dScene.Faces(L_nRunF).D3DTextureIndex = -1) Then
                                        '.SetTexture 0, G_dScene.Faces(L_nRunF).D3DTextureObject
                                        .SetRenderState D3DRENDERSTATE_TEXTUREHANDLE, G_dScene.Textures(G_dScene.Faces(L_nRunF).D3DTextureIndex).D3DHandle
                                        
                                    Else
                                        '.SetTexture 0, Nothing
                                        .SetRenderState D3DRENDERSTATE_TEXTUREHANDLE, 0
                                    End If

                                ' Enable/disable translucency
                                    If G_dUser.DisplayOptions.Translucent Then
                                        .SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, IIf(G_dScene.Faces(L_nRunF).Translucent, 1, 0)
                                    Else
                                        .SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, 0
                                    End If

                                 ' Enable/disable transparency
                                    If G_dUser.DisplayOptions.Transparent Then
                                        .SetRenderState D3DRENDERSTATE_COLORKEYENABLE, 1
                                    Else
                                        .SetRenderState D3DRENDERSTATE_COLORKEYENABLE, 0
                                    End If
                                    
                            ' Draw group triangles...
                                .Begin D3DPT_TRIANGLELIST, D3DFVF_VERTEX, 0
                                For L_nRunV = 0 To G_dScene.Faces(L_nRunF).D3DDataCount - 1
                                    .Vertex G_dScene.Faces(L_nRunF).D3DData(L_nRunV)
                                Next
                                .End 0
                            
                        End If
                    Next
                    
                    ' End scene
                    .EndScene
                    
                End With
            
            ' Draw HUD display
                
                With L_dRenderArea
                    .Top = 0
                    .Left = G_dUser.LookH
                    .Bottom = 20
                    .Right = .Left + 120
                End With
                G_oDDBackBuffer.BltFast 260, G_dRenderArea.Top + 5, G_oDDCompassSurface, L_dRenderArea, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                
                
            ' Redraw primary ...
            
                If G_dDXSelectedDriver.DriverType = EDXDTPlus Then
                    ' Flip DirectDraw buffers by hardware pageflipping
                    G_oDDPrimary.Flip Nothing, DDFLIP_WAIT
                Else
                    ' Flip DirectDraw buffers by blitting backbuffer to primary ...
                    G_oDDPrimary.BltFast G_dRenderArea.Left, G_dRenderArea.Top, G_oDDBackBuffer, G_dRenderArea, DDBLTFAST_NOCOLORKEY Or DDBLTFAST_WAIT
                End If
                
            ' Various updating operations ...
                            
                ' Update scrolling text ...
                                    
                    ' Reset texture
                    Set G_dScene.Textures(1).D3DObject = Nothing
                    
                    ' Render new text clip onto texture surface
                    With L_dRenderArea
                       .Top = G_nFrameCount Mod 480
                       .Bottom = IIf(.Top > 352, 480, .Top + 128)
                       .Left = 0
                       .Right = 128
                    End With
                    G_dScene.Textures(1).DDSurface.BltFast 0, 0, G_oDDTextSurface, L_dRenderArea, DDBLTFAST_NOCOLORKEY Or DDBLTFAST_WAIT
                    
                    ' Render upper text clip to lower area if at end
                    If G_nFrameCount Mod 480 > 352 Then
                        With L_dRenderArea
                           .Top = 0
                           .Bottom = (G_nFrameCount Mod 480) - 352
                           .Left = 0
                           .Right = 128
                        End With
                        G_dScene.Textures(1).DDSurface.BltFast 0, 128 - (G_nFrameCount Mod 480 - 352), G_oDDTextSurface, L_dRenderArea, DDBLTFAST_NOCOLORKEY Or DDBLTFAST_WAIT
                    End If
                    
                   ' Set texture
                   Set G_dScene.Textures(1).D3DObject = G_dScene.Textures(1).DDSurface
                   G_dScene.Textures(1).D3DObject.GetHandle G_oD3DDevice, G_dScene.Textures(1).D3DHandle
                   
                ' Update water ...

                    ' Reset texture
                    Set G_dScene.Textures(4).D3DObject = Nothing
                    
                    ' Render new water clip onto texture surface
                    With L_dRenderArea
                        .Top = 32 + Sin((G_nFrameCount Mod 360) * PIFactor) * 30
                        .Bottom = .Top + 64
                        .Left = 32 + Cos((G_nFrameCount Mod 360) * PIFactor) * 10
                        .Right = .Left + 64
                    End With
                    AdvancedBlit 0, 0, G_dScene.Textures(4).DDSurface, G_oDDWaterSurface, L_dRenderArea
                    
                    ' Set texture
                    Set G_dScene.Textures(4).D3DObject = G_dScene.Textures(4).DDSurface
                    G_dScene.Textures(4).D3DObject.GetHandle G_oD3DDevice, G_dScene.Textures(4).D3DHandle
                      
                ' Update flame ...

                    ' Reset texture
                    Set G_dScene.Textures(8).D3DObject = Nothing

                    ' Render new flame clip onto texture surface
                    With L_dRenderArea
                        .Top = ((G_nFrameCount Mod 16) \ 4) * 32
                        .Bottom = .Top + 32
                        .Left = (G_nFrameCount Mod 4) * 32
                        .Right = .Left + 32
                    End With
                    G_dScene.Textures(8).DDSurface.BltFast 0, 0, G_oDDFlameSurface, L_dRenderArea, DDBLTFAST_NOCOLORKEY Or DDBLTFAST_WAIT

                    ' Set texture
                    Set G_dScene.Textures(8).D3DObject = G_dScene.Textures(8).DDSurface
                    G_dScene.Textures(8).D3DObject.GetHandle G_oD3DDevice, G_dScene.Textures(8).D3DHandle
                
                ' Update flame light ...
                    If G_nFrameCount Mod 2 = 0 Then
                       G_dScene.Lights(9).D3DData.dcvColor.r = Rnd * 0.5 + 0.2
                       G_dScene.Lights(9).D3DObject.SetLight G_dScene.Lights(9).D3DData
                    End If
                    
            ' React to user input ...
            
                DoEvents
                AppInput
                
            ' Expire frame time
                Do While L_nNextFrametime > timeGetTime
                Loop
            
        Loop Until Not G_bAppRunning
    
    ' Error handling...
    
        Exit Sub
        
E_AppLoop:
        Resume Next
        AppError Err.Number, Err.Description, "AppLoop"
        
End Sub

' APPINPUT: Processes user input
Public Sub AppInput()

    ' Enable error handling ...
    On Error GoTo E_AppInput
        
    ' Setup local variables
        Dim L_nNewX As Single           ' New user X coordinates
        Dim L_nNewZ As Single           ' New user Z coordinates
        Dim L_nOldAlt As Single         ' Old altitude
        Dim L_nNewAlt As Single         ' New Altitude
        Dim L_nAltitudeChange As Single ' Amount of change in altitude
    
    ' Process user keyboad input
        With G_dUser.InputState
        
            Select Case .KeyCode
            
                ' End application
                Case vbKeyEscape
                    AppTerminate
                    G_bAppRunning = False
                    G_dUser.InputState.KeyCode = 0
                    
                ' Move forward
                Case vbKeyUp
                    L_nNewX = G_dUser.Position.X + G_dUser.Speed * Cos(G_dUser.LookH * PIFactor)
                    L_nNewZ = G_dUser.Position.z + G_dUser.Speed * Sin(G_dUser.LookH * PIFactor)
                
                ' Move backwards
                Case vbKeyDown
                    L_nNewX = G_dUser.Position.X - G_dUser.Speed * Cos(G_dUser.LookH * PIFactor)
                    L_nNewZ = G_dUser.Position.z - G_dUser.Speed * Sin(G_dUser.LookH * PIFactor)
                
                ' Step left
                Case vbKeyLeft
                    L_nNewX = G_dUser.Position.X + G_dUser.Speed * Cos(IIf(G_dUser.LookH - 90 < 0, G_dUser.LookH + 270, G_dUser.LookH - 90) * PIFactor)
                    L_nNewZ = G_dUser.Position.z + G_dUser.Speed * Sin(IIf(G_dUser.LookH - 90 < 0, G_dUser.LookH + 270, G_dUser.LookH - 90) * PIFactor)
                    
                ' Step right
                Case vbKeyRight
                    L_nNewX = G_dUser.Position.X + G_dUser.Speed * Cos(IIf(G_dUser.LookH + 90 > 359, G_dUser.LookH - 270, G_dUser.LookH + 90) * PIFactor)
                    L_nNewZ = G_dUser.Position.z + G_dUser.Speed * Sin(IIf(G_dUser.LookH + 90 > 359, G_dUser.LookH - 270, G_dUser.LookH + 90) * PIFactor)
                                
                ' Increase viewport size
                Case vbKeyAdd
                    If G_dUser.DisplaySize > 10 Then
                        G_dUser.DisplaySize = G_dUser.DisplaySize - 10
                        ViewportInitialize G_nDisplayWidth - G_dUser.DisplaySize, G_nDisplayHeight - Int(G_dUser.DisplaySize * 0.75)
                    End If
            
                ' Decrease viewport size
                Case vbKeySubtract
                    If G_dUser.DisplaySize < 380 Then
                        G_dUser.DisplaySize = G_dUser.DisplaySize + 10
                        ViewportInitialize G_nDisplayWidth - G_dUser.DisplaySize, G_nDisplayHeight - Int(G_dUser.DisplaySize * 0.75)
                    End If
            End Select
            
        End With
        
    ' Process user mouse input
        With G_dUser.InputState
        
            ' Turn head to right
            If .MouseX > 320 Then
                G_dUser.LookH = G_dUser.LookH + 2
                If G_dUser.LookH > 359 Then G_dUser.LookH = G_dUser.LookH - 360
            End If
            
            ' Turn head to left
            If .MouseX < 320 Then
                G_dUser.LookH = G_dUser.LookH - 2
                If G_dUser.LookH < 0 Then G_dUser.LookH = G_dUser.LookH + 360
            End If
            
            ' Look up
            If .MouseY > 240 Then
                G_dUser.LookV = G_dUser.LookV + 2
                If G_dUser.LookV > 60 Then G_dUser.LookV = 60
            End If
            
            ' Look down
            If .MouseY < 240 Then
                G_dUser.LookV = G_dUser.LookV - 2
                If G_dUser.LookV < -60 Then G_dUser.LookV = -60
            End If
            
            ' Reset mouse position
            SetCursorPos 320, 240
            
        End With
        
        
    ' React to position changes ...
        
        ' Check altitude change
        L_nOldAlt = G_dScene.Terrain(Int(G_dUser.Position.X), Int(G_dUser.Position.z))
        L_nNewAlt = G_dScene.Terrain(Int(L_nNewX), Int(L_nNewZ))
        L_nAltitudeChange = Abs(L_nOldAlt - L_nNewAlt)
                
        ' Check for obstacle
        If L_nAltitudeChange > 1 Then
            If Int(G_dUser.Position.X) <> Int(L_nNewX) Then
                L_nNewX = G_dUser.Position.X
            End If
            If Int(G_dUser.Position.z) <> Int(L_nNewZ) Then
                L_nNewZ = G_dUser.Position.z
            End If
        End If
        
        ' Play sounds if necessary
        If (Int(G_dUser.Position.X) <> Int(L_nNewX) Or Int(G_dUser.Position.z) <> Int(L_nNewZ)) Then
        
            ' Play metallic step sound
            If L_nNewX >= 60 And L_nNewX <= 100 And L_nNewZ >= 40 And L_nNewZ <= 45 Then
                G_oDSBStepHard.Stop
                G_oDSBStepHard.Play ByVal 0&, ByVal 0&, 0
            ElseIf L_nNewX >= 110 And L_nNewX <= 115 And L_nNewZ >= 50 And L_nNewZ <= 80 Then
                G_oDSBStepHard.Stop
                G_oDSBStepHard.Play ByVal 0&, ByVal 0&, 0
                
            ' Play soft step sound
            Else
                G_oDSBStepSoft.Stop
                G_oDSBStepSoft.Play ByVal 0&, ByVal 0&, 0
            End If
            
        End If
        
        ' Set new position
        G_dUser.Position.X = L_nNewX
        G_dUser.Position.z = L_nNewZ
        
        ' Adjust altitude if necessary
        If L_nAltitudeChange = 1 Then
            G_dUser.Position.Y = 44 - G_dScene.Terrain(Int(G_dUser.Position.X), Int(G_dUser.Position.z))
        End If
        
    ' Error handling

        Exit Sub
    
E_AppInput:

        AppError 0, "General error", "AppInput"
    
End Sub

' APPERROR: Reports application errors and terminates application properly
Public Sub AppError(nNumber As Long, sText As String, sSource As String)

    ' Enable error handling
    On Error GoTo E_AppError
    
    ' Cleanup
    Call AppTerminate
    
    ' Display error
    fMsg.Hide
    fMsg.lblTitle = "EYE3D encountered an error!"
    fMsg.lblText = IIf(InStr(1, UCase(sText), "AUTOM") > 0, "DirectX reports '" & GetDXError(nNumber) & "'", " Application reports '" & sText & "'") & vbCrLf & "SOURCE: " & sSource
    fApp.Hide
    fMsg.Show 1
    
    ' Terminate program
    End
    
    ' Error handling ...
        
        Exit Sub
        
E_AppError:

        Resume Next
    
End Sub

' APPDRIVERDETECT: Detects best DD driver, fills array of possible D3D drivers
Public Function AppDriverDetect()

    ' Enable error handling...
        On Error GoTo E_AppDriverDetect
    
    ' Setup local variables...
        Dim L_oDDInstance As IDirectDraw4   ' DD Instance for checking
        Dim L_oD3DInstance As IDirect3D3    ' D3D Instance for checking
        
    ' Detect DD driver ...
    
        ' Error handling during enumeration
        On Error Resume Next
            
        ' Enumerate directdraw drivers
        G_bPrimaryDisplayAlreadyDetected = False
        DirectDrawEnumerate AddressOf EnumDDDeviceCallback, 0
    
        ' Initialize driver types
        G_dDXDriverSoft.DriverType = EDXDTSoft
        G_dDXDriverHard.DriverType = EDXDTHard
        G_dDXDriverPlus.DriverType = EDXDTPlus
        
        ' Fetch enumeration errors
        If Err.Number > 0 Then
            AppError 0, "Error during detection of DirectDraw drivers", "AppDriverDetect"
            Exit Function
        End If
        
        ' Reset error handling
        On Error GoTo E_AppDriverDetect
        
        ' Check if at least primary driver found
        If Not G_bPrimaryDisplayAlreadyDetected Then
            AppError 0, "No valid DirectDraw driver found", "AppDriverDetect"
            Exit Function
        End If
            
    ' Detect D3D drivers ...
        
        ' Detect software driver on primary display device ...
        
            ' Create instance of DirectDraw using driver found
            DirectDrawCreate G_dDXDriverSoft.GUID, L_oDDInstance, Nothing
            
            ' Check instance existance, terminate if missing
            If L_oDDInstance Is Nothing Then
                AppError 0, "Unable to create DirectDraw instance using detected driver", "AppDriverDetect"
                Exit Function
            End If
            
            ' Look for software 3D driver
            G_dDXSelectedDriver.DriverType = EDXDTSoft
            
            ' Query DirectDraw for D3D interface
            Set L_oD3DInstance = L_oDDInstance
        
            ' Check instance existance, terminate if missing
            If L_oD3DInstance Is Nothing Then
               AppError 0, "DirectDraw interface did not return valid Direct3D interface", "AppDriverDetect"
               Exit Function
            End If

            ' Error handling during enumeration
            On Error Resume Next
            
            ' Enumerate Direct3D drivers
            L_oD3DInstance.EnumDevices AddressOf EnumD3DDeviceCallback, 0
        
            ' Catch any error resulting from the enumeration and terminate
            If Err.Number > 0 Then
                AppError 0, "Error during detection of Direct3D drivers", "AppDriverDetect"
                Exit Function
            End If
            
            ' Reset error handling
            On Error GoTo E_AppDriverDetect
        
            ' Cleanup
            Set L_oD3DInstance = Nothing
            Set L_oDDInstance = Nothing
            
            
        ' Detect hardware driver on primary display device ...
            
            ' Create instance of DirectDraw using driver found
            DirectDrawCreate G_dDXDriverSoft.GUID, L_oDDInstance, Nothing
            
            ' Check instance existance, terminate if missing
            If L_oDDInstance Is Nothing Then
                AppError 0, "Unable to create DirectDraw instance using detected driver", "AppDriverDetect"
                Exit Function
            End If
            
            ' Look for software 3D driver
            G_dDXSelectedDriver.DriverType = EDXDTHard
            
            ' Query DirectDraw for D3D interface
            Set L_oD3DInstance = L_oDDInstance
        
            ' Check instance existance, terminate if missing
            If L_oD3DInstance Is Nothing Then
               AppError 0, "DirectDraw interface did not return valid Direct3D interface", "AppDriverDetect"
               Exit Function
            End If

            ' Error handling during enumeration
            On Error Resume Next
            
            ' Enumerate Direct3D drivers
            L_oD3DInstance.EnumDevices AddressOf EnumD3DDeviceCallback, 0
        
            ' Catch any error resulting from the enumeration and terminate
            If Err.Number > 0 Then
                AppError 0, "Error during detection of Direct3D drivers", "AppDriverDetect"
                Exit Function
            End If
            
            ' Reset error handling
            On Error GoTo E_AppDriverDetect
        
            ' Cleanup
            Set L_oD3DInstance = Nothing
            Set L_oDDInstance = Nothing
            
        ' Detect hardware driver on addon board ...
        
            If G_dDXDriverPlus.Found Then
            
                ' Okay, DD driver found, but now we have to look for a D3D driver (perhaps not installed properly)!
                G_dDXDriverPlus.Found = True
            
                ' Create instance of DirectDraw using driver found
                DirectDrawCreate G_dDXDriverPlus.GUID, L_oDDInstance, Nothing
                
                ' Check instance existance, terminate if missing
                If L_oDDInstance Is Nothing Then
                    AppError 0, "Unable to create DirectDraw instance using detected driver", "AppDriverDetect"
                    Exit Function
                End If
                
                ' Look for software 3D driver
                G_dDXSelectedDriver.DriverType = EDXDTPlus
                
                ' Query DirectDraw for D3D interface
                Set L_oD3DInstance = L_oDDInstance
            
                ' Check instance existance, terminate if missing
                If L_oD3DInstance Is Nothing Then
                   AppError 0, "DirectDraw interface did not return valid Direct3D interface", "AppDriverDetect"
                   Exit Function
                End If
    
                ' Error handling during enumeration
                On Error Resume Next
                
                ' Enumerate Direct3D drivers
                L_oD3DInstance.EnumDevices AddressOf EnumD3DDeviceCallback, 0
            
                ' Catch any error resulting from the enumeration and terminate
                If Err.Number > 0 Then
                    AppError 0, "Error during detection of Direct3D drivers", "AppDriverDetect"
                    Exit Function
                End If
                
                ' Reset error handling
                On Error GoTo E_AppDriverDetect
            
                ' Cleanup
                Set L_oD3DInstance = Nothing
                Set L_oDDInstance = Nothing
                
            End If
                
    ' Set selected driver
        If G_dDXDriverPlus.Found Then
            G_dDXSelectedDriver = G_dDXDriverPlus
        ElseIf G_dDXDriverHard.Found Then
            G_dDXSelectedDriver = G_dDXDriverHard
        Else
            G_dDXSelectedDriver = G_dDXDriverSoft
        End If
        
    ' Error handling
        Exit Function
        
E_AppDriverDetect:
        
    ' Cleanup...
            
        On Error Resume Next
        
        ' Release interfaces
        Set L_oD3DInstance = Nothing
        Set L_oDDInstance = Nothing
    
    ' Error report
    
        AppError 0, "General error during driver detection", "AppDriverDetect"
    
End Function

' VIEWPORTINITIALIZE: Initializes a given D3DIM viewport to passed size
Public Sub ViewportInitialize(nWidth As Integer, nHeight As Integer)

    ' Enable error handling ...
    On Error GoTo E_ViewportInitialize
    
    ' Setup local variables ...
        
        Dim L_dD3DViewportDesc As D3DVIEWPORT2      ' Description of viewport object for generation of viewport
        Dim L_dDDBLTFX As DDBLTFX                   ' FX Blit descriptor
        Dim L_dRenderArea As RECT                   ' Rectangle for clearing whole backbuffer
        
    ' Setup viewport ...
    
        ' Fill viewport description
        With L_dD3DViewportDesc
            .dwSize = Len(L_dD3DViewportDesc)
            .dwX = (G_nDisplayWidth - nWidth) / 2
            .dwY = (G_nDisplayHeight - nHeight) / 2
            .dwWidth = nWidth
            .dwHeight = nHeight
            .dvClipX = -1
            .dvClipY = 1
            .dvClipHeight = 2
            .dvClipWidth = 2
            .dvMinZ = 0
            .dvMaxZ = 1
        End With
        
        ' Set viewport properties
        G_oD3DViewport.SetViewport2 L_dD3DViewportDesc

        ' Setup render area for rendering loop
        With G_dRenderArea
            .Top = (G_nDisplayHeight - nHeight) / 2
            .Left = (G_nDisplayWidth - nWidth) / 2
            .Right = .Left + nWidth
            .Bottom = .Top + nHeight
        End With
        With G_dClearArea
            .X1 = 0
            .Y1 = 0
            .X2 = G_nDisplayWidth
            .Y2 = G_nDisplayHeight
        End With
        
    ' Clear buffer
        With L_dDDBLTFX
            .dwSize = Len(L_dDDBLTFX)
            .dwFillColor = 0
        End With
        G_oDDBackBuffer.Blt G_dClearArea, ByVal Nothing, ByVal 0&, DDBLT_COLORFILL Or DDBLT_WAIT, L_dDDBLTFX
        G_oDDPrimary.Blt G_dClearArea, ByVal Nothing, ByVal 0&, DDBLT_COLORFILL Or DDBLT_WAIT, L_dDDBLTFX
        
    ' Error handling ...
        
        Exit Sub
        
E_ViewportInitialize:

    AppError Err.Number, Err.Description, "ViewportInitialize"

End Sub

' SCENEINITIALIZE: Loads 3D data from text file and sets up environment
Public Sub SceneInitialize()

    ' Enable Error handling ...
        
        On Error GoTo E_SceneInitialize
    
    ' Setup local variables ...
        
        Dim L_nRunStar As Integer               ' Variable to run through all stars
        Dim L_nStarColor As Integer             ' Current star color
        Dim L_nRunX As Integer                  ' Variable to run through X coordinates
        Dim L_nRunY As Integer                  ' Variable to run through X coordinates
        Dim L_sInString As String               ' String to read from file
        Dim L_nRunV As Integer                  ' Variable to run through all vertices
        Dim L_nMaterialIndex As Integer         ' Index of current material
        Dim L_nTransformIndex As Integer        ' Index of current transform group
        Dim L_nTranslucent As Integer           ' Translucency flag
        Dim L_nItemCount As Integer             ' Counter for input from file
        Dim L_nRunItem As Integer               ' Variable to run through file input
        Dim L_nTransparent As Integer           ' Transparency flagg
        Dim L_dDDCK As DDCOLORKEY               ' Color key for transparency
        
    ' Load scene data ...
            
        ' Open scene data file
        Open App.Path + "\scene.dat" For Input As #1
        
        ' Input lights ...
        
            ' Get light count, size light array
            Input #1, L_nItemCount
            ReDim G_dScene.Lights(L_nItemCount)
            
            ' Load all lights
            For L_nRunItem = 0 To L_nItemCount - 1
                            
                ' Read light data
                With G_dScene.Lights(L_nRunItem).D3DData
                    .dwSize = Len(G_dScene.Lights(L_nRunItem).D3DData)
                    Input #1, .dltType, .dcvColor.r, .dcvColor.g, .dcvColor.b, .dcvColor.a, .dvRange, .dvPosition.X, .dvPosition.Y, .dvPosition.z, .dvDirection.X, .dvDirection.Y, .dvDirection.z, .dvPhi, .dvTheta, .dvFalloff, .dvAttenuation0, .dvAttenuation1, .dvAttenuation2, .dwFlags
                End With

                ' Create light object, add it to viewport
                With G_dScene.Lights(L_nRunItem)
                    G_oD3DInstance.CreateLight .D3DObject, Nothing
                    .D3DObject.SetLight .D3DData
                    G_oD3DViewport.AddLight .D3DObject
                End With
                
            Next
            
        ' Input textures ...
            
            ' Get texture count, resize texture array
            Input #1, L_nItemCount
            ReDim G_dScene.Textures(L_nItemCount)
            
            ' Load all textures
            For L_nRunItem = 0 To L_nItemCount - 1
                With G_dScene.Textures(L_nRunItem)
                    
                    ' Read texture data
                    Input #1, .Filename, .Width, .Height, L_nTransparent
                    .Transparent = IIf(L_nTransparent = 1, True, False)
                    
                    
                    ' Create/Load texture objects
                    If .Filename = "NONE" Then
                        Set .DDSurface = CreateTexture(.Width)
                    Else
                        Set .DDSurface = LoadTexture(App.Path + "\" + .Filename)
                    End If
                    
                    ' Set color key if transparency enabled
                    If .Transparent Then
                        L_dDDCK.dwColorSpaceLowValue = 0
                        L_dDDCK.dwColorSpaceHighValue = 0
                        .DDSurface.SetColorKey DDCKEY_SRCBLT, L_dDDCK
                    End If
                    
                    ' Get texture handle
                    Set .D3DObject = .DDSurface
                    .D3DObject.GetHandle G_oD3DDevice, .D3DHandle
                    
                End With
            Next
                    
        ' Input materials ...
        
            ' Get material count, size material array
            Input #1, L_nItemCount
            ReDim G_dScene.Materials(L_nItemCount)
            
            ' Load all materials
            For L_nRunItem = 0 To L_nItemCount - 1
                            
                ' Read material data
                With G_dScene.Materials(L_nRunItem).D3DData
                    .dwSize = Len(G_dScene.Materials(L_nRunItem).D3DData)
                    Input #1, G_dScene.Materials(L_nRunItem).D3DTextureIndex, .Ambient.r, .Ambient.g, .Ambient.b, .Ambient.a, .diffuse.r, .diffuse.g, .diffuse.b, .diffuse.a, .Specular.r, .Specular.g, .Specular.b, .Specular.a, .emissive.r, .emissive.g, .emissive.b, .emissive.a, .power
                End With

                ' Create material object
                With G_dScene.Materials(L_nRunItem)
                    G_oD3DInstance.CreateMaterial .D3DObject, Nothing
                    .D3DObject.SetMaterial .D3DData
                    .D3DObject.GetHandle G_oD3DDevice, .D3DHandle
                End With
                
            Next
        
        ' Input meshdata ...
            ReDim G_dScene.Faces(19)
            
            ' Go through records
            Do While Not EOF(1)
                
                ' Read index of transform group and material
                Input #1, L_nTransformIndex, L_nMaterialIndex, L_nItemCount, L_nTranslucent
                
                ' Set properties of face group
                With G_dScene.Faces(L_nTransformIndex)
                    ' Set face group to enabled
                    .Enabled = True
                    ' Set face group translucency status
                    .Translucent = IIf(L_nTranslucent = 1, True, False)
                    ' Set face group vertex count
                    .D3DDataCount = L_nItemCount * 3
                    ' Reference material of group
                    .D3DMaterialIndex = L_nMaterialIndex
                    ' Reference texture of group
                    .D3DTextureIndex = G_dScene.Materials(.D3DMaterialIndex).D3DTextureIndex
                    ' Reference material
                    MIdentity .D3DTransform
                End With
                
                ' Read in vertices
                For L_nRunV = 0 To L_nItemCount - 1
                    With G_dScene.Faces(L_nTransformIndex)
                        Input #1, .D3DData(L_nRunV * 3).X, .D3DData(L_nRunV * 3).Y, .D3DData(L_nRunV * 3).z, .D3DData(L_nRunV * 3).nx, .D3DData(L_nRunV * 3).ny, .D3DData(L_nRunV * 3).nz, .D3DData(L_nRunV * 3).tu, .D3DData(L_nRunV * 3).tv, .D3DData(L_nRunV * 3 + 1).X, .D3DData(L_nRunV * 3 + 1).Y, .D3DData(L_nRunV * 3 + 1).z, .D3DData(L_nRunV * 3 + 1).nx, .D3DData(L_nRunV * 3 + 1).ny, .D3DData(L_nRunV * 3 + 1).nz, .D3DData(L_nRunV * 3 + 1).tu, .D3DData(L_nRunV * 3 + 1).tv, .D3DData(L_nRunV * 3 + 2).X, .D3DData(L_nRunV * 3 + 2).Y, .D3DData(L_nRunV * 3 + 2).z, .D3DData(L_nRunV * 3 + 2).nx, .D3DData(L_nRunV * 3 + 2).ny, .D3DData(L_nRunV * 3 + 2).nz, .D3DData(L_nRunV * 3 + 2).tu, .D3DData(L_nRunV * 3 + 2).tv
                    End With
                Next
                
            Loop
            
        ' Finish...
        
        Close #1
        
    ' Load terrain altitude data
    
        ' Open file ...
        Open App.Path + "\terrain.dat" For Input As #1
        
        ' Read data ...
        For L_nRunX = 0 To 149
            Input #1, L_sInString
            For L_nRunY = 0 To 149
                G_dScene.Terrain(L_nRunX, L_nRunY) = Val(Mid(L_sInString, L_nRunY + 1, 1))
            Next
        Next
        
        ' Finish ...
        Close #1
        
    
    ' Load miscellaneous data (bitmaps)
    
        ' Load text bitmap
        Set G_oDDTextSurface = LoadSurface(App.Path + "\texturetext.bmp")
        
        ' Load water bitmap
        Set G_oDDWaterSurface = LoadSurface(App.Path + "\texturewater.bmp")
        
        ' Load flame bitmap
        Set G_oDDFlameSurface = LoadSurface(App.Path + "\textureflame.bmp")
        
        ' Create color key for compass
        With L_dDDCK
            .dwColorSpaceHighValue = 0
            .dwColorSpaceLowValue = 0
        End With
        
        ' Load compass bitmap
        Set G_oDDCompassSurface = LoadSurface(App.Path + "\texturecompass.bmp")
        G_oDDCompassSurface.SetColorKey DDCKEY_SRCBLT, L_dDDCK
        
    ' Setup star-sprenkled sky (cylinder-projection of 2d-points) ...
    
        For L_nRunStar = 0 To 1999
            With G_dScene.Stars(L_nRunStar)
                .Altitude = Int(Rnd * 1000) - 500
                .Direction = Int(Rnd * 3600)
                L_nStarColor = 125 - .Altitude \ 5 + IIf(Rnd > 0.66, Rnd * 100 - 50, 0)
                If L_nStarColor < 10 Then L_nStarColor = 10
                If L_nStarColor > 250 Then L_nStarColor = 250
                .Color = RGB(L_nStarColor, L_nStarColor, L_nStarColor)
            End With
        Next
        
    ' Initialize 3D sound ...
        
        ' Load wave into sound buffer
        Set G_oDSBDisplaySound = LoadWaveAudio(App.Path + "\Display.wav", True)
        
        ' Create 3D sound buffer, set properties
        Set G_oDS3DBDisplaySound = G_oDSBDisplaySound
        With G_oDS3DBDisplaySound
            .SetMinDistance 1, DS3D_IMMEDIATE
            .SetMaxDistance 50, DS3D_IMMEDIATE
            .SetMode DS3DMODE_NORMAL, DS3D_IMMEDIATE
            .SetPosition 35, 48, 50, DS3D_IMMEDIATE         '(Display)
        End With
        
        ' Start playing sound
        G_oDSBDisplaySound.Play ByVal 0&, ByVal 0&, DSBPLAY_LOOPING
            
    ' Initialize other sound ...
    
        ' Load wave into sound buffer
        Set G_oDSBStepHard = LoadWaveAudio(App.Path + "\stephard.wav")
        Set G_oDSBStepSoft = LoadWaveAudio(App.Path + "\stepsoft.wav")

    ' Error handling ...
        
        Exit Sub
        
E_SceneInitialize:

    Close #1
    AppError Err.Number, Err.Description, "SceneInitialize"
    
End Sub

' SCENETERMINATE: Releases all 3D data
Public Sub SceneTerminate()

    ' Enable Error handling ...
        
        On Error Resume Next
    
    ' Setup local variables ...
        
        Dim nRun As Integer             ' Variable to run through various arrays
        Dim nMaterialIndex As Integer   ' Current Material index
        Dim nTransformIndex As Integer  ' Index of current transform group
        Dim nVertexCount As Integer     ' Number of vertices
        
    ' Remove lights ...
        
        For nRun = 0 To UBound(G_dScene.Lights)
            Set G_dScene.Lights(nRun).D3DObject = Nothing
        Next
    
    ' Remove materials ...
        
        For nRun = 0 To UBound(G_dScene.Materials)
            Set G_dScene.Materials(nRun).D3DObject = Nothing
        Next
        
    ' Remove textures  ...
        
        For nRun = 0 To UBound(G_dScene.Textures)
            Set G_dScene.Textures(nRun).D3DObject = Nothing
            Set G_dScene.Textures(nRun).DDSurface = Nothing
        Next
    
    ' Remove miscellaneos ...
        Set G_oDDTextSurface = Nothing
        Set G_oDDWaterSurface = Nothing
        Set G_oDDFlameSurface = Nothing
        Set G_oDDCompassSurface = Nothing
        
    ' Remove sounds ...
       
        G_oDSBDisplaySound.Stop
        Set G_oDS3DBDisplaySound = Nothing
        Set G_oDSBDisplaySound = Nothing
    
        G_oDSBStepSoft.Stop
        Set G_oDSBStepSoft = Nothing
        
        G_oDSBStepHard.Stop
        Set G_oDSBStepHard = Nothing
        
    ' Error handling ...
        
        On Error GoTo 0
    
End Sub
