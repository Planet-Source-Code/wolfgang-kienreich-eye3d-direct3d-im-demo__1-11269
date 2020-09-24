Attribute VB_Name = "mDefs"
Option Explicit

    
' Constants for use with win32 API ...
    
    Public Const IMAGE_BITMAP = 0
    Public Const LR_LOADFROMFILE = &H10
    Public Const LR_CREATEDIBSECTION = &H2000
    Public Const SRCCOPY = &HCC0020
    
' Constants for use with DirectX tlb
    Public Const DDSD_TEXTURESTAGE = 1048576
    Public Const DDSCAPS2_TEXTUREMANAGE = 16

' Various constants ...
    Public Const PIValue = 3.141593
    Public Const PIFactor = 0.017453
    
' Types for use with win32 API ...

    ' Wave format type
    Type WAVEFORMATEX
        wFormatTag As Integer
        nChannels As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
        wBitsPerSample As Integer
        cbSize As Integer
    End Type
    
    ' Rectangle type
    Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
    
    ' Bitmap descriptor type
    Public Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
    End Type

' Functions for use with win32 API ...

    ' Single Pixel manipulation
    Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
    Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
    
    ' DC manipulation
    Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
    Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
    Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
    Public Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
    
    ' General GDI Object manipulation
    Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
    Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
    
    ' Bitmap manipulation
    Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Public Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
    Public Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
    
    ' Various functions
    Public Declare Function timeGetTime Lib "winmm.dll" () As Long
    Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal source As Long, ByVal length As Long)
    Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
    Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

' Types for use with the application
        
    ' User type: Contains user data ...
        Public Type tOptions
            Transparent As Boolean                  ' Use texture transparency
            Translucent As Boolean                  ' Use texture translucency
            Specular As Boolean                     ' Use specular highlights
            Correct As Boolean                      ' Use perspectivic correction
            Mapping As Boolean                      ' Use texture mapping
            Phong As Boolean                        ' Use phong shading mode
            Filtering As Boolean                    ' Filter bilinear
        End Type
        Public Type tInputState
            MouseX As Single                        ' Position of mouse
            MouseY As Single                        ' Position of mouse
            MouseButton As Integer                  ' Mousebutton pressed
            KeyCode As Long                         ' Key pressed
        End Type
        Public Type tStats
            Frametime As Single                     ' Average frame time
        End Type
        Public Type tUser
            Position As D3DVECTOR                      ' Position of user within world
            LookH As Integer                           ' Direction user is facing
            LookV As Integer                           ' Elevation of user look
            Speed As Single                            ' Movement speed of user
            DisplaySize As Integer                     ' Size decrease of display
            DisplayOptions As tOptions                 ' Options for displaying 3D world
            InputState As tInputState                  ' Current state of input devices
            Stats As tStats                            ' Statistics
        End Type
    
    ' Direct3D Scene data types ...
        Public Type tLight
            D3DData As D3DLIGHT2                       ' Light data settings
            D3DObject As IDirect3DLight                ' Actual light object
        End Type
        Public Type tTexture
            DDSurface As IDirectDrawSurface2           ' DDraw surface holding texture
            D3DObject As IDirect3DTexture2             ' Actual Texture object
            D3DHandle As Long                          ' Texture handle
            Transparent As Boolean                     ' Should a color key be created for the texture
            Filename As String                         ' Filename of texture
            Width As Integer                           ' Size of texture
            Height As Integer                          ' Size of texture
        End Type
        Public Type tMaterial
            D3DData As D3DMATERIAL                     ' Material data settings
            D3DObject As IDirect3DMaterial2            ' Actual Material object
            D3DHandle As Long                          ' Handle to material
            D3DTextureIndex As Long                    ' Reference to texture used
        End Type
        Public Type tFace
            D3DDataCount As Integer                   ' Count of vertexdata
            D3DData(1999) As D3DVERTEX                ' Vertexdata for primitive drawing using this settings; Declared static because REDIM doesn't support nested dynamic arrays
            D3DMaterialIndex As Long                  ' Handle to material to use
            D3DTextureIndex As Long                    ' Reference to texture used
            D3DTransform As D3DMATRIX                 ' Transform matrix to apply
            Enabled As Boolean                        ' Tells if this face group is in use
            Translucent As Boolean                    ' Use translucency ?
        End Type
        Public Type tStar
            Altitude As Integer                       ' Altitude of star above ground
            Direction As Integer                      ' Direction of star relative to zero
            Color As Long                             ' Star color
        End Type
        Public Type tScene
            Lights() As tLight                        ' Collection of lights
            Textures() As tTexture                    ' Collection of textures
            Materials() As tMaterial                  ' Collection of materials
            Faces() As tFace                          ' Collection of faces
            Terrain(149, 149) As Byte                 ' Terrain altitude definition
            Stars(1999) As tStar                       ' Stars to fill sky
        End Type
    
' Public Variables for use with the application ...

    ' DirectX data ...
    
        ' DirectX instance variables
        Public G_oDDInstance As IDirectDraw4                ' Instance of DirectDraw interface
        Public G_oD3DInstance As IDirect3D3                 ' Instance of Direct3DIM interface
        Public G_oDSInstance As IDirectSound                ' Instance of DirectSound interface
        
        ' DirectX display system
        Public G_oDDPrimary As IDirectDrawSurface4          ' Primary surface
        Public G_oDDBackBuffer As IDirectDrawSurface4       ' Backbuffer surface
        Public G_dRenderArea As RECT                        ' Rectangle defining output area for DirectDraw
        Public G_dClearArea As D3DRECT                      ' Rectangle definition for clearing of backbuffer
        
        ' Driver variables and arrays for driver detection
        Public G_dDXDriverHard As tDDDriver                 ' Hardware DirectDraw driver
        Public G_dDXDriverSoft As tDDDriver                 ' Software DirectDraw driver
        Public G_dDXDriverPlus As tDDDriver                 ' Accellerator add on DirectDraw driver
        Public G_dDXSelectedDriver As tDDDriver             ' Selected driver
        Public G_bPrimaryDisplayAlreadyDetected As Boolean  ' Flag used for display driver enum
        
        ' Direct3D framework
        Public G_oD3DDevice As IDirect3DDevice3             ' D3DIM device
        Public G_oD3DViewport As IDirect3DViewport3         ' D3DIM viewport
        Public G_dD3DViewportArea As D3DRECT                ' Rectangle defining viewport
    
    ' Application: General application data ...
        
        ' Frame counter
        Public G_nFrameCount As Long                        ' Frame counter
        
        ' Scene data ...
        Public G_dUser As tUser                             ' User data
        Public G_dScene As tScene                           ' D3D scene data
        Public G_oDDTextSurface As IDirectDrawSurface4      ' Surface to hold scrolling text
        Public G_oDDWaterSurface As IDirectDrawSurface4     ' Surface to hold flowing water
        Public G_oDDFlameSurface As IDirectDrawSurface4     ' Surface to hold flame animation phases
        Public G_oDDCompassSurface As IDirectDrawSurface4   ' Surface to hold compass image
        
        Public G_oDSBPrimary As IDirectSoundBuffer          ' Primary sound buffer
        Public G_oDSListener As IDirectSound3DListener      ' Listener to D3DSound
        Public G_oDSBDisplaySound As IDirectSoundBuffer     ' DirectSound Buffer for display sound
        Public G_oDS3DBDisplaySound As IDirectSound3DBuffer ' DirectSound 3D Buffer for display sound
        Public G_oDSBStepSoft As IDirectSoundBuffer         ' DirectSound buffer for step noises
        Public G_oDSBStepHard As IDirectSoundBuffer         ' DirectSound buffer for step noises
        
        ' Miscellaneous data ...
        Public G_bAppInitialized As Boolean                 ' Initialization flag
        Public G_bAppRunning As Boolean                     ' Execution flag
        
        Public G_nDisplayWidth As Integer                   ' Width of display
        Public G_nDisplayHeight As Integer                  ' Height of display
        
    
