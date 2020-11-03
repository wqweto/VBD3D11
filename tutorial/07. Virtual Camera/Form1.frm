VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6924
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9924
   LinkTopic       =   "Form1"
   ScaleHeight     =   6924
   ScaleWidth      =   9924
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z

#Const DebugBuild = True

Private Const ERROR_FILE_NOT_FOUND                      As Long = 2
Private Const LNG_FACILITY_WIN32                        As Long = &H80070000

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As D3D11_RECT) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'--- GDI+
Private Declare Function GdiplusStartup Lib "gdiplus" (hToken As Long, pInputBuf As Any, Optional ByVal pOutputBuf As Long = 0) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal sFilename As Long, hImage As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hBitmap As Long, lpRect As Any, ByVal lFlags As Long, ByVal lPixelFormat As Long, uLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hBitmap As Long, uLockedBitmapData As BitmapData) As Long

Private Type BitmapData
    Width               As Long
    Height              As Long
    Stride              As Long
    PixelFormat         As Long
    Scan0               As Long
    Reserved            As Long
End Type

Private Enum UcsGameAction
    GameActionMoveCamFwd
    GameActionMoveCamBack
    GameActionMoveCamLeft
    GameActionMoveCamRight
    GameActionTurnCamLeft
    GameActionTurnCamRight
    GameActionLookUp
    GameActionLookDown
    GameActionRaiseCam
    GameActionLowerCam
    GameActionCount
End Enum

Private m_d3d11Device           As ID3D11Device1
Private m_d3d11DeviceContext    As ID3D11DeviceContext1
Private m_d3d11SwapChain        As IDXGISwapChain1
Private m_d3d11FrameBufferView  As ID3D11RenderTargetView
Private m_vsBlob                As ID3DBlob
Private m_vertexShader          As ID3D11VertexShader
Private m_pixelShader           As ID3D11PixelShader
Private m_inputLayout           As ID3D11InputLayout
Private m_vertexBuffer          As ID3D11Buffer
Private m_numVerts              As Long
Private m_stride                As Long
Private m_offset                As Long
Private m_samplerState          As ID3D11SamplerState
Private m_textureView           As ID3D11ShaderResourceView
Private m_constantBuffer        As ID3D11Buffer
Private m_rasterizerState       As ID3D11RasterizerState
Private m_cameraPos             As XMFLOAT3
Private m_cameraFwd             As XMFLOAT3
Private m_cameraPitch           As Single
Private m_cameraYaw             As Single
Private m_currentTimeInSeconds  As Double
Private m_isRunning             As Boolean
Private m_windowDidResize       As Boolean
Private m_keyIsDown(0 To GameActionCount - 1) As Boolean

Private Type UcsBuffer
    Data()              As Byte
End Type

Private Type UcsConstants
    modelViewProj       As XMMATRIX
End Type
Private Const sizeof_Single         As Long = 4
Private Const sizeof_UcsConstants   As Long = 16 * sizeof_Single

Private Sub pvHandleKey(KeyCode As Integer, Shift As Integer, ByVal bDown As Boolean)
    #If Shift Then
    #End If
    Select Case KeyCode
    Case vbKeyEscape
        m_isRunning = False
    Case vbKeyW
        m_keyIsDown(GameActionMoveCamFwd) = bDown
    Case vbKeyA
        m_keyIsDown(GameActionMoveCamLeft) = bDown
    Case vbKeyS
        m_keyIsDown(GameActionMoveCamBack) = bDown
    Case vbKeyD
        m_keyIsDown(GameActionMoveCamRight) = bDown
    Case vbKeyE
        m_keyIsDown(GameActionRaiseCam) = bDown
    Case vbKeyQ
        m_keyIsDown(GameActionLowerCam) = bDown
    Case vbKeyUp
        m_keyIsDown(GameActionLookUp) = bDown
    Case vbKeyLeft
        m_keyIsDown(GameActionTurnCamLeft) = bDown
    Case vbKeyDown
        m_keyIsDown(GameActionLookDown) = bDown
    Case vbKeyRight
        m_keyIsDown(GameActionTurnCamRight) = bDown
    End Select
End Sub

Private Function pvLoadPng(sFilename As String, lWidth As Long, lHeight As Long, lChannels As Long, baData() As Byte) As Boolean
    Const ImageLockModeRead As Long = 1
    Const PixelFormat32bppPARGB As Long = &HE200B
    Dim aInput(0 To 3)  As Long
    Dim hBitmap         As Long
    Dim uData           As BitmapData
    
    If GetModuleHandle("gdiplus") = 0 Then
        aInput(0) = 1
        Call GdiplusStartup(0, aInput(0))
    End If
    If GdipLoadImageFromFile(StrPtr(sFilename), hBitmap) <> 0 Then
        GoTo QH
    End If
    If GdipBitmapLockBits(hBitmap, ByVal 0, ImageLockModeRead, PixelFormat32bppPARGB, uData) <> 0 Then
        GoTo QH
    End If
    lWidth = uData.Width
    lHeight = uData.Height
    lChannels = 4
    ReDim baData(0 To uData.Stride * uData.Height - 1) As Byte
    Call CopyMemory(baData(0), ByVal uData.Scan0, UBound(baData) + 1)
    '--- success
    pvLoadPng = True
QH:
    If uData.Scan0 <> 0 Then
        Call GdipBitmapUnlockBits(hBitmap, uData)
    End If
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
    End If
End Function

Private Sub Form_Load()
    Dim hResult         As VBHRESULT
    
    '--- Create D3D11 Device and Context
    Dim featureLevels() As Long
    Dim creationFlags   As Long
    pvArrayLong featureLevels, D3D_FEATURE_LEVEL_11_0, D3D_FEATURE_LEVEL_10_1, D3D_FEATURE_LEVEL_10_0
    creationFlags = D3D11_CREATE_DEVICE_BGRA_SUPPORT
    #If DebugBuild Then
        creationFlags = creationFlags Or D3D11_CREATE_DEVICE_DEBUG
    #End If
RetryCreateDevice:
    hResult = D3D11CreateDevice(Nothing, D3D_DRIVER_TYPE_HARDWARE, 0, creationFlags, _
                                featureLevels(0), UBound(featureLevels) + 1, D3D11_SDK_VERSION, _
                                m_d3d11Device, 0, m_d3d11DeviceContext)
    If hResult = DXGI_ERROR_SDK_COMPONENT_MISSING And (creationFlags And D3D11_CREATE_DEVICE_DEBUG) <> 0 Then
        creationFlags = creationFlags And Not D3D11_CREATE_DEVICE_DEBUG
        GoTo RetryCreateDevice
    End If
    If hResult < 0 Then
        Err.Raise hResult, "D3D11CreateDevice"
    End If
    
#If DebugBuild Then
    '--- Set up debug layer to break on D3D11 errors
    Dim d3dDebug        As ID3D11Debug
    Dim d3dInfoQueue    As ID3D11InfoQueue
    Set d3dDebug = m_d3d11Device
    If Not d3dDebug Is Nothing Then
        Set d3dInfoQueue = d3dDebug
        d3dInfoQueue.SetBreakOnSeverity D3D11_MESSAGE_SEVERITY_CORRUPTION, 1
        d3dInfoQueue.SetBreakOnSeverity D3D11_MESSAGE_SEVERITY_ERROR, 1
        Set d3dInfoQueue = Nothing
    End If
    Set d3dDebug = Nothing
#End If
    
    '--- Create Swap Chain
    Dim dxgiFactory     As IDXGIFactory2
    Dim dxgiDevice      As IDXGIDevice1
    Dim dxgiAdapter     As IDXGIAdapter
    Dim adapterDesc     As DXGI_ADAPTER_DESC
    Dim d3d11SwapChainDesc As DXGI_SWAP_CHAIN_DESC1
    Set dxgiDevice = m_d3d11Device
    Set dxgiAdapter = dxgiDevice.GetAdapter()
    Set dxgiDevice = Nothing
    dxgiAdapter.GetDesc adapterDesc
    Debug.Print "Graphics Device: " & Replace(adapterDesc.Description, vbNullChar, vbNullString)
    Set dxgiFactory = dxgiAdapter.GetParent(IIDFromString(szIID_IDXGIFactory2))
    With d3d11SwapChainDesc
        .Width = 0 '--- use window width
        .Height = 0 '--- use window height
        .Format = DXGI_FORMAT_B8G8R8A8_UNORM_SRGB
        .SampleDesc.Count = 1
        .SampleDesc.Quality = 0
        .BufferUsage = DXGI_USAGE_RENDER_TARGET_OUTPUT
        .BufferCount = 2
        .Scaling = DXGI_SCALING_STRETCH
        .SwapEffect = DXGI_SWAP_EFFECT_DISCARD
        .AlphaMode = DXGI_ALPHA_MODE_UNSPECIFIED
        .Flags = 0
    End With
    Set m_d3d11SwapChain = dxgiFactory.CreateSwapChainForHwnd(m_d3d11Device, hWnd, d3d11SwapChainDesc, ByVal 0, Nothing)

    '--- Create Framebuffer Render Target
    Dim d3d11FrameBuffer As ID3D11Texture2D
    Set d3d11FrameBuffer = m_d3d11SwapChain.GetBuffer(0, IIDFromString(szIID_ID3D11Texture2D))
    Set m_d3d11FrameBufferView = m_d3d11Device.CreateRenderTargetView(d3d11FrameBuffer, ByVal 0)
    Set d3d11FrameBuffer = Nothing
    
    '--- Create Vertex Shader
    Dim shaderCompileErrorsBlob As ID3DBlob
    Dim errorString     As String
    hResult = D3DCompileFromFile(PathCombine(App.Path, "shaders.hlsl"), ByVal 0, ByVal 0, "vs_main", "vs_4_0", 0, 0, m_vsBlob, shaderCompileErrorsBlob)
    If hResult < 0 Then
        If hResult = LNG_FACILITY_WIN32 Or ERROR_FILE_NOT_FOUND Then
            errorString = "Could not compile shader; file not found"
        ElseIf Not shaderCompileErrorsBlob Is Nothing Then
            errorString = pvToString(shaderCompileErrorsBlob.GetBufferPointer())
        End If
        MsgBox errorString, vbCritical, "Shader Compiler Error"
        Unload Me
        GoTo QH
    End If
    Set m_vertexShader = m_d3d11Device.CreateVertexShader(m_vsBlob.GetBufferPointer(), m_vsBlob.GetBufferSize(), Nothing)
    
    '--- Create Pixel Shader
    Dim psBlob As ID3DBlob
    hResult = D3DCompileFromFile(PathCombine(App.Path, "shaders.hlsl"), ByVal 0, ByVal 0, "ps_main", "ps_4_0", 0, 0, psBlob, shaderCompileErrorsBlob)
    If hResult < 0 Then
        If hResult = LNG_FACILITY_WIN32 Or ERROR_FILE_NOT_FOUND Then
            errorString = "Could not compile shader; file not found"
        ElseIf Not shaderCompileErrorsBlob Is Nothing Then
            errorString = pvToString(shaderCompileErrorsBlob.GetBufferPointer())
        End If
        MsgBox errorString, vbCritical, "Shader Compiler Error"
        Unload Me
        GoTo QH
    End If
    Set m_pixelShader = m_d3d11Device.CreatePixelShader(psBlob.GetBufferPointer(), psBlob.GetBufferSize(), Nothing)
    Set psBlob = Nothing
    
    '--- Create Input Layout
    Dim inputElementDesc(0 To 1)  As D3D11_INPUT_ELEMENT_DESC
    Dim nameBuffer(0 To 1) As UcsBuffer
    pvInitInputElementDesc inputElementDesc(0), nameBuffer(0), "POS", 0, DXGI_FORMAT_R32G32_FLOAT, 0, 0, D3D11_INPUT_PER_VERTEX_DATA, 0
    pvInitInputElementDesc inputElementDesc(1), nameBuffer(1), "TEX", 0, DXGI_FORMAT_R32G32_FLOAT, 0, D3D11_APPEND_ALIGNED_ELEMENT, D3D11_INPUT_PER_VERTEX_DATA, 0
    Set m_inputLayout = m_d3d11Device.CreateInputLayout(inputElementDesc(0), UBound(inputElementDesc) + 1, m_vsBlob.GetBufferPointer(), m_vsBlob.GetBufferSize())
    
    '--- Create Vertex Buffer
    Dim vertexData() As Single '--- x, y, u, v
    pvArraySingle vertexData, _
        -0.5!, 0.5!, 0!, 0!, _
        0.5!, -0.5!, 1!, 1!, _
        -0.5!, -0.5!, 0!, 1!, _
        -0.5!, 0.5!, 0!, 0!, _
        0.5!, 0.5!, 1!, 0!, _
        0.5!, -0.5!, 1!, 1!
    m_stride = 4 * sizeof_Single
    m_numVerts = (UBound(vertexData) + 1) * sizeof_Single / m_stride
    m_offset = 0
    Dim vertexBufferDesc As D3D11_BUFFER_DESC
    With vertexBufferDesc
        .ByteWidth = (UBound(vertexData) + 1) * sizeof_Single
        .Usage = D3D11_USAGE_IMMUTABLE
        .BindFlags = D3D11_BIND_VERTEX_BUFFER
    End With
    Dim vertexSubresourceData As D3D11_SUBRESOURCE_DATA
    vertexSubresourceData.pSysMem = VarPtr(vertexData(0))
    Set m_vertexBuffer = m_d3d11Device.CreateBuffer(vertexBufferDesc, vertexSubresourceData)
    
    '--- Create Sampler State
    Dim samplerDesc As D3D11_SAMPLER_DESC
    With samplerDesc
        .Filter = D3D11_FILTER_MIN_MAG_MIP_POINT
        .AddressU = D3D11_TEXTURE_ADDRESS_BORDER
        .AddressV = D3D11_TEXTURE_ADDRESS_BORDER
        .AddressW = D3D11_TEXTURE_ADDRESS_BORDER
        .BorderColor(0) = 1!
        .BorderColor(1) = 1!
        .BorderColor(2) = 1!
        .BorderColor(3) = 1!
        .ComparisonFunc = D3D11_COMPARISON_NEVER
    End With
    Set m_samplerState = m_d3d11Device.CreateSamplerState(samplerDesc)
    
    '--- Load Image
    Dim texWidth         As Long
    Dim texHeight        As Long
    Dim texNumChannels   As Long
    Dim testTextureBytes() As Byte
    Dim texBytesPerRow   As Long
    If Not pvLoadPng(PathCombine(App.Path, "testTexture.png"), texWidth, texHeight, texNumChannels, testTextureBytes) Then
        MsgBox "Error loading testTexture.png", vbExclamation, "Load Image"
        Unload Me
        GoTo QH
    End If
    texBytesPerRow = texWidth * texNumChannels
    
    '--- Create Texture
    Dim textureDesc     As D3D11_TEXTURE2D_DESC
    Dim textureSubresourceData As D3D11_SUBRESOURCE_DATA
    Dim texture         As ID3D11Texture2D
    With textureDesc
        .Width = texWidth
        .Height = texHeight
        .MipLevels = 1
        .ArraySize = 1
        .Format = DXGI_FORMAT_B8G8R8A8_UNORM_SRGB
        .SampleDesc.Count = 1
        .Usage = D3D11_USAGE_IMMUTABLE
        .BindFlags = D3D11_BIND_SHADER_RESOURCE
    End With
    With textureSubresourceData
        .pSysMem = VarPtr(testTextureBytes(0))
        .SysMemPitch = texBytesPerRow
    End With
    Set texture = m_d3d11Device.CreateTexture2D(textureDesc, textureSubresourceData)
    Set m_textureView = m_d3d11Device.CreateShaderResourceView(texture, ByVal 0)
    
    '--- Create Constant Buffer
    Dim constantBufferDesc As D3D11_BUFFER_DESC
    With constantBufferDesc
        .ByteWidth = (sizeof_UcsConstants + &HF&) And &HFFFFFFF0 '--- ByteWidth must be a multiple of 16, per the docs
        .Usage = D3D11_USAGE_DYNAMIC
        .BindFlags = D3D11_BIND_CONSTANT_BUFFER
        .CPUAccessFlags = D3D11_CPU_ACCESS_WRITE
    End With
    Set m_constantBuffer = m_d3d11Device.CreateBuffer(constantBufferDesc, ByVal 0)
    
    Dim rasterizerDesc As D3D11_RASTERIZER_DESC
    With rasterizerDesc
        .FillMode = D3D11_FILL_SOLID
        .CullMode = D3D11_CULL_NONE
        .FrontCounterClockwise = 1
    End With
    Set m_rasterizerState = m_d3d11Device.CreateRasterizerState(rasterizerDesc)
    
    '--- Camera
    Dim perspectiveMat As XMMATRIX
    m_cameraPos = XmMake3(0, 0, 2)
    m_cameraFwd = XmMake3(0, 0, -1)
    m_cameraPitch = 0!
    m_cameraYaw = 0!
    
    '--- Timing
    Dim startTime As Double
    startTime = TimerEx
    
    '--- Main Loop
    Show
    m_isRunning = True
    Do While m_isRunning
        Dim dt              As Single
        Dim previousTimeInSeconds As Double
        previousTimeInSeconds = m_currentTimeInSeconds
        m_currentTimeInSeconds = TimerEx - startTime
        dt = m_currentTimeInSeconds - previousTimeInSeconds
        If dt > 1! / 60! Then
            dt = 1! / 60!
        End If
        Caption = "[" & Format$(m_currentTimeInSeconds, "0.000") & " - " & Format$(dt, "0.000") & "]"
        
        '--- Get window dimensions
        Dim windowWidth As Long
        Dim windowHeight As Long
        Dim windowAspectRatio As Single
        Dim clientRect As D3D11_RECT
        Call GetClientRect(hWnd, clientRect)
        windowWidth = clientRect.Right - clientRect.Left
        windowHeight = clientRect.Bottom - clientRect.Top
        windowAspectRatio = CSng(windowWidth) / CSng(windowHeight)
        
        If m_windowDidResize Then
            m_d3d11DeviceContext.OMSetRenderTargets 0, Nothing, Nothing
            Set m_d3d11FrameBufferView = Nothing
            m_d3d11SwapChain.ResizeBuffers 0, 0, 0, DXGI_FORMAT_UNKNOWN, 0
            Set d3d11FrameBuffer = m_d3d11SwapChain.GetBuffer(0, IIDFromString(szIID_ID3D11Texture2D))
            Set m_d3d11FrameBufferView = m_d3d11Device.CreateRenderTargetView(d3d11FrameBuffer, ByVal 0)
            Set d3d11FrameBuffer = Nothing
            
            perspectiveMat = XmMakePerspectiveMat(windowAspectRatio, DegreesToRadians(84), 0.1!, 1000!)
            
            m_windowDidResize = False
        End If
        
        '--- Update camera
        Dim camFwdXZ As XMFLOAT3
        Dim cameraRightXZ As XMFLOAT3
        camFwdXZ = XmNormalize(XmMake3(m_cameraFwd.x, 0, m_cameraFwd.z))
        cameraRightXZ = XmCross(camFwdXZ, XmMake3(0, 1, 0))

        Const CAM_MOVE_SPEED As Single = 5!  '--- in metres per second
        Dim CAM_MOVE_AMOUNT As Single
        CAM_MOVE_AMOUNT = CAM_MOVE_SPEED * dt
        If m_keyIsDown(GameActionMoveCamFwd) Then
            m_cameraPos = XmAdd(m_cameraPos, XmScalarMul(camFwdXZ, CAM_MOVE_AMOUNT))
        End If
        If m_keyIsDown(GameActionMoveCamBack) Then
            m_cameraPos = XmSub(m_cameraPos, XmScalarMul(camFwdXZ, CAM_MOVE_AMOUNT))
        End If
        If m_keyIsDown(GameActionMoveCamLeft) Then
            m_cameraPos = XmSub(m_cameraPos, XmScalarMul(cameraRightXZ, CAM_MOVE_AMOUNT))
        End If
        If m_keyIsDown(GameActionMoveCamRight) Then
            m_cameraPos = XmAdd(m_cameraPos, XmScalarMul(cameraRightXZ, CAM_MOVE_AMOUNT))
        End If
        If m_keyIsDown(GameActionRaiseCam) Then
            m_cameraPos.y = m_cameraPos.y + CAM_MOVE_AMOUNT
        End If
        If m_keyIsDown(GameActionLowerCam) Then
            m_cameraPos.y = m_cameraPos.y + CAM_MOVE_AMOUNT
        End If
        
        Const CAM_TURN_SPEED As Single = M_PI '--- in radians per second
        Dim CAM_TURN_AMOUNT As Single
        CAM_TURN_AMOUNT = CAM_TURN_SPEED * dt
        If m_keyIsDown(GameActionTurnCamLeft) Then
            m_cameraYaw = m_cameraYaw + CAM_TURN_AMOUNT
        End If
        If m_keyIsDown(GameActionTurnCamRight) Then
            m_cameraYaw = m_cameraYaw - CAM_TURN_AMOUNT
        End If
        If m_keyIsDown(GameActionLookUp) Then
            m_cameraPitch = m_cameraPitch + CAM_TURN_AMOUNT
        End If
        If m_keyIsDown(GameActionLookDown) Then
            m_cameraPitch = m_cameraPitch - CAM_TURN_AMOUNT
        End If

        '--- Wrap yaw to avoid floating-point errors if we turn too far
        Do While m_cameraYaw >= 2 * M_PI
            m_cameraYaw = m_cameraYaw - 2 * M_PI
        Loop
        Do While m_cameraYaw <= -2 * M_PI
            m_cameraYaw = m_cameraYaw + 2 * M_PI
        Loop

        '--- Clamp pitch to stop camera flipping upside down
        If m_cameraPitch > DegreesToRadians(85) Then
            m_cameraPitch = DegreesToRadians(85)
        End If
        If m_cameraPitch < -DegreesToRadians(85) Then
            m_cameraPitch = -DegreesToRadians(85)
        End If
        
        '--- Calculate view matrix from camera data
        '
        ' float4x4 viewMat = inverse(rotateXMat(cameraPitch) * rotateYMat(cameraYaw) * translationMat(cameraPos));
        ' NOTE: We can simplify this calculation to avoid inverse()!
        ' Applying the rule inverse(A*B) = inverse(B) * inverse(A) gives:
        ' float4x4 viewMat = inverse(translationMat(cameraPos)) * inverse(rotateYMat(cameraYaw)) * inverse(rotateXMat(cameraPitch));
        ' The inverse of a rotation/translation is a negated rotation/translation:
        Dim viewMat As XMMATRIX
        viewMat = XmMulMat(XmMulMat( _
            XmTranslationMat(XmNeg(m_cameraPos)), _
            XmRotateYMat(-m_cameraYaw)), _
            XmRotateXMat(-m_cameraPitch))
        
        '--- Update the forward vector we use for camera movement:
        m_cameraFwd = XmMake3(viewMat.m(2, 0), viewMat.m(2, 1), -viewMat.m(2, 2))

        '--- Spin the quad
        Dim modelMat As XMMATRIX
        modelMat = XmRotateYMat(0.2! * (M_PI * m_currentTimeInSeconds))
        
        '--- Calculate model-view-projection matrix to send to shader
        Dim modelViewProj As XMMATRIX
        modelViewProj = XmMulMat(XmMulMat( _
            modelMat, _
            viewMat), _
            perspectiveMat)
        
        '--- Update constant buffer
        Dim mappedSubresource As D3D11_MAPPED_SUBRESOURCE
        m_d3d11DeviceContext.Map m_constantBuffer, 0, D3D11_MAP_WRITE_DISCARD, 0, mappedSubresource
        Dim constants As UcsConstants
        constants.modelViewProj = modelViewProj
        Call CopyMemory(ByVal mappedSubresource.pData, constants, sizeof_UcsConstants)
        m_d3d11DeviceContext.Unmap m_constantBuffer, 0
        
        Dim backgroundColor() As Single
        pvArraySingle backgroundColor, 0.1!, 0.2!, 0.6!, 1!
        m_d3d11DeviceContext.ClearRenderTargetView m_d3d11FrameBufferView, backgroundColor(0)
        
        Dim viewport As D3D11_VIEWPORT
        pvInitViewport viewport, 0, 0, windowWidth, windowHeight, 0, 1
        m_d3d11DeviceContext.RSSetViewports 1, viewport
        m_d3d11DeviceContext.RSSetState m_rasterizerState
        
        m_d3d11DeviceContext.OMSetRenderTargets 1, m_d3d11FrameBufferView, Nothing
        
        m_d3d11DeviceContext.IASetPrimitiveTopology D3D11_PRIMITIVE_TOPOLOGY_TRIANGLELIST
        m_d3d11DeviceContext.IASetInputLayout m_inputLayout
        
        m_d3d11DeviceContext.VSSetShader m_vertexShader, Nothing, 0
        m_d3d11DeviceContext.PSSetShader m_pixelShader, Nothing, 0
        
        m_d3d11DeviceContext.PSSetShaderResources 0, 1, m_textureView
        m_d3d11DeviceContext.PSSetSamplers 0, 1, m_samplerState
        
        m_d3d11DeviceContext.VSSetConstantBuffers 0, 1, m_constantBuffer
        
        m_d3d11DeviceContext.IASetVertexBuffers 0, 1, m_vertexBuffer, m_stride, m_offset

        m_d3d11DeviceContext.Draw m_numVerts, 0
        
        m_d3d11SwapChain.Present 1, 0
        
        DoEvents
    Loop
QH:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_isRunning = False
End Sub

Private Sub Form_Resize()
    m_windowDidResize = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    pvHandleKey KeyCode, Shift, True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    pvHandleKey KeyCode, Shift, False
End Sub

'=========================================================================
' Shared
'=========================================================================

Private Sub pvArrayLong(aDest() As Long, ParamArray a() As Variant)
    Dim lIdx            As Long
    
    ReDim aDest(0 To UBound(a)) As Long
    For lIdx = 0 To UBound(a)
        aDest(lIdx) = a(lIdx)
    Next
End Sub

Private Sub pvArraySingle(aDest() As Single, ParamArray a() As Variant)
    Dim lIdx            As Long
    
    ReDim aDest(0 To UBound(a)) As Single
    For lIdx = 0 To UBound(a)
        aDest(lIdx) = a(lIdx)
    Next
End Sub

Private Sub pvInitInputElementDesc(uEntry As D3D11_INPUT_ELEMENT_DESC, uBuffer As UcsBuffer, SemanticName As String, ByVal SemanticIndex As Long, ByVal Format As DXGI_FORMAT, ByVal InputSlot As Long, ByVal AlignedByteOffset As Long, ByVal InputSlotClass As D3D11_INPUT_CLASSIFICATION, ByVal InstanceDataStepRate As Long)
    uBuffer.Data = StrConv(SemanticName & vbNullChar, vbFromUnicode)
    With uEntry
        .SemanticName = VarPtr(uBuffer.Data(0))
        .SemanticIndex = SemanticIndex
        .Format = Format
        .InputSlot = InputSlot
        .AlignedByteOffset = AlignedByteOffset
        .InputSlotClass = InputSlotClass
        .InstanceDataStepRate = InstanceDataStepRate
    End With
End Sub

Private Sub pvInitViewport(uEntry As D3D11_VIEWPORT, ByVal TopLeftX As Single, ByVal TopLeftY As Single, ByVal Width As Single, ByVal Height As Single, ByVal MinDepth As Single, ByVal MaxDepth As Single)
    With uEntry
        .TopLeftX = TopLeftX
        .TopLeftY = TopLeftY
        .Width = Width
        .Height = Height
        .MinDepth = MinDepth
        .MaxDepth = MaxDepth
    End With
End Sub

Private Function pvToString(ByVal lPtr As Long) As String
    If lPtr <> 0 Then
        pvToString = String$(lstrlen(lPtr), 0)
        Call CopyMemory(ByVal pvToString, ByVal lPtr, Len(pvToString))
    End If
End Function

Private Function PathCombine(sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\" And LenB(sFile) <> 0, "\", vbNullString) & sFile
End Function

Private Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency
    
    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property
