Attribute VB_Name = "DirectXInit"
Option Explicit

' These Main Components of Direct.
Public DX               As New DirectX7         'DirectX
Public DD               As DirectDraw4          'DirectDraw
Public D3D              As Direct3DRM3          'Direct3D

' DirectDraw Surfaces - Where the screen is draw.
Public Primary          As DirectDrawSurface4   'Primary Surface
Public BackBuffer       As DirectDrawSurface4   'BackBuffer

' ViewPort and Direct3D Device
Public D3D_Device       As Direct3DRMDevice3    'Mode Device Direct3D
Public D3D_ViewPort     As Direct3DRMViewport2  'Viewport Direct3D

' Frames
Public FrameRoot         As Direct3DRMFrame3
Public FrameLight        As Direct3DRMFrame3
Public FrameCamera       As Direct3DRMFrame3

Public Ddsd1            As DDSURFACEDESC2       'DirectDraw Surface Description

Public DI_State         As DIKEYBOARDSTATE      'Direct Input State KeyBoard
Public DI_Device        As DirectInputDevice    'Direct Input Device

Public xCamera          As Long                 'Pos Camera in x
Public yCamera          As Long                 'Pos Camera in y
Public zCamera          As Long                 'Pos Camera in z

' Picture Buffer
Public PicBuffer As DirectDrawSurface4
Public PicBufferRECT As RECT

' Frame count/set
Public StartTick As Long
Public LastTick As Long
Public NowTime As Long

Public Function D3DInit(ByRef NameForm As Form, ByVal ScrWidth As Long, ByVal ScrHeight As Long, ByVal ScrColor As Integer)
    
    On Error Resume Next
'-------------------------------------------------------------------
    Set DD = DX.DirectDraw4Create("")
    NameForm.Show
'-------------------------------------------------------------------
    Call DD.SetCooperativeLevel(NameForm.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE) 'Set Screen Mode in Full Screen
    Call DD.SetDisplayMode(ScrWidth, ScrHeight, ScrColor, 0, DDSDM_DEFAULT)         'Set Resolution and BitDepth
'-------------------------------------------------------------------
    ' Fill out primary surface description
    Ddsd1.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    Ddsd1.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_3DDEVICE Or DDSCAPS_COMPLEX Or DDSCAPS_FLIP
    Ddsd1.lBackBufferCount = 1
    
    ' Primary is the primary surface
    Set Primary = DD.CreateSurface(Ddsd1)
'-------------------------------------------------------------------
    ' Get the backbuffer
    Dim Caps As DDSCAPS2
    Caps.lCaps = DDSCAPS_BACKBUFFER
    Set BackBuffer = Primary.GetAttachedSurface(Caps)
    BackBuffer.SetForeColor RGB(255, 255, 0)    ' Set SetForeColor to Yellow
'-------------------------------------------------------------------
    ' Direct 3D Initializes
    Set D3D = DX.Direct3DRMCreate()             ' Creates the Direct3D
    ' using device Direct3D, hardware rendering (HALDevice) or software enumeration (RGBDevice).
    Set D3D_Device = D3D.CreateDeviceFromSurface("IID_IDirect3DHALDevice", DD, BackBuffer, D3DRMDEVICE_DEFAULT)
    D3D_Device.SetBufferCount 2                 ' Set the number of buffers
    D3D_Device.SetQuality D3DRMRENDER_GOURAUD   ' Rendering Quality, GOURAUD (best rendering quality).
'-------------------------------------------------------------------
    D3D_Device.SetTextureQuality D3DRMTEXTURE_LINEAR             ' Set Texture Quality
    D3D_Device.SetRenderMode D3DRMRENDERMODE_BLENDEDTRANSPARENCY ' Set Render Mode.

    D3D_Device.SetDither D_TRUE
'-------------------------------------------------------------------
    ' Initialize KeyBoard and Mouse
    Dim DI_Main     As DirectInput

    Set DI_Main = DX.DirectInputCreate()
    Set DI_Device = DI_Main.CreateDevice("GUID_SysKeyboard")
    DI_Device.SetCommonDataFormat DIFORMAT_KEYBOARD
    DI_Device.SetCooperativeLevel NameForm.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    DI_Device.Acquire
    
End Function

' Initializes Frame for Root, Camera and Light
Sub FrameD3DInit(ByVal ScrWidth As Long, ByVal ScrHeight As Long)
    
    ' Lights
    Dim LightAmbient    As Direct3DRMLight
    Dim LightSpot       As Direct3DRMLight
    
    Set FrameRoot = D3D.CreateFrame(Nothing)
    Set FrameCamera = D3D.CreateFrame(FrameRoot)
    Set FrameLight = D3D.CreateFrame(FrameCamera)
'-------------------------------------------------------------------
    ' Set the BackGround color a Black.
    FrameRoot.SetSceneBackgroundRGB 0, 0, 0
'-------------------------------------------------------------------
    ' Set Camera equivalen with screen Width x Height pixels
    xCamera = ScrWidth / 2                  ' set Camera position in Direct3D to equivalen
    yCamera = -(ScrHeight / 2)              ' with screen 800 x 600 pixels
    zCamera = -ScrWidth                     ' add - 500 to test view create Map
    ' Set Camera position
    FrameCamera.SetPosition Nothing, xCamera, yCamera, zCamera
    Set D3D_ViewPort = D3D.CreateViewport(D3D_Device, FrameCamera, 0, 0, ScrWidth, ScrHeight)
'-------------------------------------------------------------------
    ' How far see meshes object (*.x), to now use arrow Up / Down after Run
    D3D_ViewPort.SetBack ScrWidth + 1500     ' zCamera > 1300 meshes Object will not Draw
'-------------------------------------------------------------------
    ' Create & Set Point Light, Type, Color
    FrameLight.SetPosition Nothing, 1, 6, -20
    Set LightSpot = D3D.CreateLightRGB(D3DRMLIGHT_POINT, 2, 2, 2)
    FrameLight.AddLight LightSpot
'-------------------------------------------------------------------
    ' Create & Set Ambient Light
    Set LightAmbient = D3D.CreateLightRGB(D3DRMLIGHT_AMBIENT, 1, 1, 1)
    FrameRoot.AddLight LightAmbient
    
End Sub

Public Sub LoadBMPandSurface(surface As DirectDrawSurface4, BmpPath As String, RECTvar As RECT, Optional TransparantColor As Integer = 0)
    Dim Tempddsd As DDSURFACEDESC2  ' Store to Temp DirectDraw surface Description
    
    Set surface = Nothing

    ' Load image sprite
    Tempddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    Tempddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Set surface = DD.CreateSurfaceFromFile(BmpPath, Tempddsd)

'-------------------------------------------------------------------
    Dim ColourKey As DDCOLORKEY     ' Colourkey to make Transparant or Not
    ColourKey.low = TransparantColor
    ColourKey.high = TransparantColor
    surface.SetColorKey DDCKEY_SRCBLT, ColourKey
End Sub

Public Sub DelayGame(Delay As Integer)
    StartTick = DX.TickCount
    NowTime = DX.TickCount
    Do Until NowTime - LastTick > Delay
        DoEvents
        NowTime = DX.TickCount
    Loop
    LastTick = NowTime
End Sub

