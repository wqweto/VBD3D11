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

'=========================================================================
' API
'=========================================================================

Private Const ERROR_FILE_NOT_FOUND                      As Long = 2
Private Const LNG_FACILITY_WIN32                        As Long = &H80070000

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As D3D11_RECT) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

'=========================================================================
' Enums
'=========================================================================

Private Enum UcsGameActionEnum
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

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const sizeof_Single         As Long = 4
Private Const sizeof_UcsConstants   As Long = 16 * sizeof_Single
Private Const sizeof_Integer        As Long = 2

Private m_d3d11Device           As ID3D11Device1
Private m_d3d11DeviceContext    As ID3D11DeviceContext1
Private m_d3d11SwapChain        As IDXGISwapChain1
Private m_d3d11FrameBufferView  As ID3D11RenderTargetView
Private m_depthBufferView       As ID3D11DepthStencilView
Private m_vsBlob                As ID3DBlob
Private m_vertexShader          As ID3D11VertexShader
Private m_pixelShader           As ID3D11PixelShader
Private m_inputLayout           As ID3D11InputLayout
Private m_vertexBuffer          As ID3D11Buffer
Private m_indexBuffer           As ID3D11Buffer
'Private m_numVerts              As Long
Private m_numIndices            As Long
Private m_stride                As Long
Private m_offset                As Long
Private m_samplerState          As ID3D11SamplerState
Private m_textureView           As ID3D11ShaderResourceView
Private m_constantBuffer        As ID3D11Buffer
Private m_rasterizerState       As ID3D11RasterizerState
Private m_depthStencilState     As ID3D11DepthStencilState
Private m_perspectiveMat        As XMMATRIX
Private m_cameraPos             As XMFLOAT3
Private m_cameraFwd             As XMFLOAT3
Private m_cameraPitch           As Single
Private m_cameraYaw             As Single
Private m_currentTimeInSeconds  As Double
Private m_isRunning             As Boolean
Private m_windowDidResize       As Boolean
Private m_keyIsDown(0 To GameActionCount - 1) As Boolean

Private Type UcsBufferType
    Data()              As Byte
End Type

Private Type UcsConstantsType
    modelViewProj       As XMMATRIX
End Type

'=========================================================================
' Methods
'=========================================================================

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

Private Sub pvCreateD3D11RenderTargets( _
            d3d11Device As ID3D11Device1, _
            swapChain As IDXGISwapChain1, _
            d3d11FrameBufferView As ID3D11RenderTargetView, _
            depthBufferView As ID3D11DepthStencilView)
    Dim aGUID(0 To 4)   As Long
    Dim d3d11FrameBuffer As ID3D11Texture2D
    Dim depthBufferDesc As D3D11_TEXTURE2D_DESC
    Dim depthBuffer     As ID3D11Texture2D
    
    Call IIDFromString(szIID_ID3D11Texture2D, aGUID(0))
    Set d3d11FrameBuffer = swapChain.GetBuffer(0, aGUID(0))
    Set d3d11FrameBufferView = m_d3d11Device.CreateRenderTargetView(d3d11FrameBuffer, ByVal 0)
    d3d11FrameBuffer.GetDesc depthBufferDesc
    Set d3d11FrameBuffer = Nothing

    depthBufferDesc.Format = DXGI_FORMAT_D24_UNORM_S8_UINT
    depthBufferDesc.BindFlags = D3D11_BIND_DEPTH_STENCIL
    Set depthBuffer = d3d11Device.CreateTexture2D(depthBufferDesc, ByVal 0)
    Set depthBufferView = d3d11Device.CreateDepthStencilView(depthBuffer, ByVal 0)
End Sub

Private Sub pvMainLoop()
    Dim hResult         As VBHRESULT
    Dim aGUID(0 To 3)   As Long
    
    '--- Create D3D11 Device and Context
    Dim featureLevels() As Long
    Dim creationFlags   As Long
    pvArrayLong featureLevels, D3D_FEATURE_LEVEL_11_0
    creationFlags = D3D11_CREATE_DEVICE_BGRA_SUPPORT
    #If DebugBuild Then
        creationFlags = creationFlags Or D3D11_CREATE_DEVICE_DEBUG
    #End If
    hResult = D3D11CreateDevice(Nothing, D3D_DRIVER_TYPE_HARDWARE, 0, creationFlags, _
                                featureLevels(0), UBound(featureLevels) + 1, D3D11_SDK_VERSION, _
                                m_d3d11Device, 0, m_d3d11DeviceContext)
    If hResult < 0 Then
        Err.Raise hResult, "D3D11CreateDevice"
    End If
    
#If DebugBuild And False Then
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
    Call IIDFromString(szIID_IDXGIFactory2, aGUID(0))
    Set dxgiFactory = dxgiAdapter.GetParent(aGUID(0))
    With d3d11SwapChainDesc
        .Width = 0 '--- use window width
        .Height = 0 '--- use window height
        .Format = DXGI_FORMAT_B8G8R8A8_UNORM_SRGB
        .SampleDesc.Count = 2
        .SampleDesc.Quality = 0
        .BufferUsage = DXGI_USAGE_RENDER_TARGET_OUTPUT
        .BufferCount = 2
        .Scaling = DXGI_SCALING_STRETCH
        .SwapEffect = DXGI_SWAP_EFFECT_DISCARD
        .AlphaMode = DXGI_ALPHA_MODE_UNSPECIFIED
        .Flags = 0
    End With
    Set m_d3d11SwapChain = dxgiFactory.CreateSwapChainForHwnd(m_d3d11Device, hWnd, d3d11SwapChainDesc, ByVal 0, Nothing)

    '--- Create Render Target and Depth Buffer
    pvCreateD3D11RenderTargets m_d3d11Device, m_d3d11SwapChain, m_d3d11FrameBufferView, m_depthBufferView
    
    '--- Create Vertex Shader
    Dim shaderCompileErrorsBlob As ID3DBlob
    Dim errorString     As String
    Dim shaderCompileFlags As Long
    '--- Compiling with this flag allows debugging shaders with Visual Studio
    #If DebugBuild Then
        shaderCompileFlags = shaderCompileFlags Or D3DCOMPILE_DEBUG
    #End If
    hResult = D3DCompileFromFile(PathCombine(App.Path, "shaders.hlsl"), ByVal 0, ByVal 0, "vs_main", "vs_5_0", shaderCompileFlags, 0, m_vsBlob, shaderCompileErrorsBlob)
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
    hResult = D3DCompileFromFile(PathCombine(App.Path, "shaders.hlsl"), ByVal 0, ByVal 0, "ps_main", "ps_5_0", shaderCompileFlags, 0, psBlob, shaderCompileErrorsBlob)
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
    Dim inputElementDesc(0 To 0)  As D3D11_INPUT_ELEMENT_DESC
    Dim nameBuffer(0 To 0) As UcsBufferType
    pvInitInputElementDesc inputElementDesc(0), nameBuffer(0), "POS", 0, DXGI_FORMAT_R32G32B32_FLOAT, 0, 0, D3D11_INPUT_PER_VERTEX_DATA, 0
    Set m_inputLayout = m_d3d11Device.CreateInputLayout(inputElementDesc(0), UBound(inputElementDesc) + 1, m_vsBlob.GetBufferPointer(), m_vsBlob.GetBufferSize())
    
    '--- Create Vertex and Index Buffer
    Dim vertexData() As Single '--- x, y, z
    pvArraySingle vertexData, _
        -0.5!, -0.5!, -0.5!, _
        -0.5!, -0.5!, 0.5!, _
        -0.5!, 0.5!, -0.5!, _
        -0.5!, 0.5!, 0.5!, _
        0.5!, -0.5!, -0.5!, _
        0.5!, -0.5!, 0.5!, _
        0.5!, 0.5!, -0.5!, _
        0.5!, 0.5!, 0.5!
    Dim indices() As Integer
    pvArrayInteger indices, _
        0, 6, 4, _
        0, 2, 6, _
        0, 3, 2, _
        0, 1, 3, _
        2, 7, 6, _
        2, 3, 7, _
        4, 6, 7, _
        4, 7, 5, _
        0, 4, 5, _
        0, 5, 1, _
        1, 5, 7, _
        1, 7, 3
    m_stride = 3 * sizeof_Single
'    m_numVerts = (UBound(vertexData) + 1) * sizeof_Single / m_stride
    m_offset = 0
    m_numIndices = UBound(indices) + 1
    Dim vertexBufferDesc As D3D11_BUFFER_DESC
    With vertexBufferDesc
        .ByteWidth = (UBound(vertexData) + 1) * sizeof_Single
        .Usage = D3D11_USAGE_IMMUTABLE
        .BindFlags = D3D11_BIND_VERTEX_BUFFER
    End With
    Dim vertexSubresourceData As D3D11_SUBRESOURCE_DATA
    vertexSubresourceData.pSysMem = VarPtr(vertexData(0))
    Set m_vertexBuffer = m_d3d11Device.CreateBuffer(vertexBufferDesc, vertexSubresourceData)
    
    Dim indexBufferDesc  As D3D11_BUFFER_DESC
    With indexBufferDesc
        .ByteWidth = (UBound(indices) + 1) * sizeof_Integer
        .Usage = D3D11_USAGE_IMMUTABLE
        .BindFlags = D3D11_BIND_INDEX_BUFFER
    End With
    Dim indexSubresourceData As D3D11_SUBRESOURCE_DATA
    indexSubresourceData.pSysMem = VarPtr(indices(0))
    Set m_indexBuffer = m_d3d11Device.CreateBuffer(indexBufferDesc, indexSubresourceData)
    
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
        .MultisampleEnable = 1
    End With
    Set m_rasterizerState = m_d3d11Device.CreateRasterizerState(rasterizerDesc)
    
    Dim depthStencilDesc  As D3D11_DEPTH_STENCIL_DESC
    With depthStencilDesc
        .DepthEnable = 1
        .DepthWriteMask = D3D11_DEPTH_WRITE_MASK_ALL
        .DepthFunc = D3D11_COMPARISON_LESS
    End With
    Set m_depthStencilState = m_d3d11Device.CreateDepthStencilState(depthStencilDesc)
    
    '--- Camera
    m_cameraPos = XmMake3(0, 0, 2)
    m_cameraFwd = XmMake3(0, 0, -1)
    m_cameraPitch = 0!
    m_cameraYaw = 0!
    
    '--- Main Loop
    Show
    m_isRunning = True
    Do While m_isRunning
        Dim dt              As Single
        Dim previousTimeInSeconds As Double
        previousTimeInSeconds = m_currentTimeInSeconds
        m_currentTimeInSeconds = TimerEx
        dt = m_currentTimeInSeconds - previousTimeInSeconds
        If dt > 1! / 60! Then
            dt = 1! / 60!
        End If
        
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
            Set m_depthBufferView = Nothing
            m_d3d11SwapChain.ResizeBuffers 0, 0, 0, DXGI_FORMAT_UNKNOWN, 0
            
            pvCreateD3D11RenderTargets m_d3d11Device, m_d3d11SwapChain, m_d3d11FrameBufferView, m_depthBufferView
            m_perspectiveMat = XmMakePerspectiveMat(windowAspectRatio, DegreesToRadians(84), 0.1!, 1000!)
            
            m_windowDidResize = False
        End If
        
        '--- Update camera
        Dim camFwdXZ        As XMFLOAT3
        Dim cameraRightXZ   As XMFLOAT3
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

        '--- Spin the cube
        Dim modelMat As XMMATRIX
        modelMat = XmMulMat( _
            XmRotateXMat(-0.2! * (M_PI * m_currentTimeInSeconds)), _
            XmRotateYMat(0.1! * (M_PI * m_currentTimeInSeconds)))
        
        '--- Calculate model-view-projection matrix to send to shader
        Dim modelViewProj As XMMATRIX
        modelViewProj = XmMulMat(XmMulMat( _
            modelMat, _
            viewMat), _
            m_perspectiveMat)
        
        '--- Update constant buffer
        Dim mappedSubresource As D3D11_MAPPED_SUBRESOURCE
        m_d3d11DeviceContext.Map m_constantBuffer, 0, D3D11_MAP_WRITE_DISCARD, 0, mappedSubresource
        Dim constants As UcsConstantsType
        constants.modelViewProj = modelViewProj
        Call CopyMemory(ByVal mappedSubresource.pData, constants, sizeof_UcsConstants)
        m_d3d11DeviceContext.Unmap m_constantBuffer, 0
        
        Dim backgroundColor() As Single
        pvArraySingle backgroundColor, 0.1!, 0.2!, 0.6!, 1!
        m_d3d11DeviceContext.ClearRenderTargetView m_d3d11FrameBufferView, backgroundColor(0)
        m_d3d11DeviceContext.ClearDepthStencilView m_depthBufferView, D3D11_CLEAR_DEPTH, 1!, 0
        
        Dim viewport    As D3D11_VIEWPORT
        pvInitViewport viewport, 0, 0, windowWidth, windowHeight, 0, 1
        m_d3d11DeviceContext.RSSetViewports 1, viewport
        m_d3d11DeviceContext.RSSetState m_rasterizerState
        
        m_d3d11DeviceContext.OMSetDepthStencilState m_depthStencilState, 0
        m_d3d11DeviceContext.OMSetRenderTargets 1, m_d3d11FrameBufferView, m_depthBufferView
        
        m_d3d11DeviceContext.IASetPrimitiveTopology D3D11_PRIMITIVE_TOPOLOGY_TRIANGLELIST
        m_d3d11DeviceContext.IASetInputLayout m_inputLayout
        
        m_d3d11DeviceContext.VSSetShader m_vertexShader, Nothing, 0
        m_d3d11DeviceContext.PSSetShader m_pixelShader, Nothing, 0
        
        m_d3d11DeviceContext.PSSetShaderResources 0, 1, m_textureView
        m_d3d11DeviceContext.PSSetSamplers 0, 1, m_samplerState
        
        m_d3d11DeviceContext.VSSetConstantBuffers 0, 1, m_constantBuffer
        
        m_d3d11DeviceContext.IASetVertexBuffers 0, 1, m_vertexBuffer, m_stride, m_offset
        m_d3d11DeviceContext.IASetIndexBuffer m_indexBuffer, DXGI_FORMAT_R16_UINT, 0

        m_d3d11DeviceContext.DrawIndexed m_numIndices, 0, 0
        
        m_d3d11SwapChain.Present 1, 0
        
        DoEvents
    Loop
QH:
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

Private Sub pvArrayInteger(aDest() As Integer, ParamArray a() As Variant)
    Dim lIdx            As Long
    
    ReDim aDest(0 To UBound(a)) As Integer
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

Private Sub pvInitInputElementDesc(uEntry As D3D11_INPUT_ELEMENT_DESC, uBuffer As UcsBufferType, SemanticName As String, ByVal SemanticIndex As Long, ByVal Format As DXGI_FORMAT, ByVal InputSlot As Long, ByVal AlignedByteOffset As Long, ByVal InputSlotClass As D3D11_INPUT_CLASSIFICATION, ByVal InstanceDataStepRate As Long)
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

'=========================================================================
' Control events
'=========================================================================

Private Sub Form_Load()
    pvMainLoop
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
