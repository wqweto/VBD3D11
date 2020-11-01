VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const DebugBuild = True

Private Const ERROR_FILE_NOT_FOUND                      As Long = 2
Private Const LNG_FACILITY_WIN32                        As Long = &H80070000

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As D3D11_RECT) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Enum UcsGameAction
    GameActionMoveUp
    GameActionMoveDown
    GameActionMoveLeft
    GameActionMoveRight
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
Private m_constantBuffer        As ID3D11Buffer
Private m_isRunning             As Boolean
Private m_windowDidResize       As Boolean
Private m_playerPos             As D3D_FLOAT2
Private m_keyIsDown(0 To GameActionCount - 1) As Boolean

Private Type UcsBuffer
    Data()              As Byte
End Type

Private Type UcsConstants
    pos                 As D3D_FLOAT2
    paddingUnused       As D3D_FLOAT2 ' color (below) needs to be 16-byte aligned!
    color               As D3D_FLOAT4
End Type
Private Const sizeof_UcsConstants As Long = 32

Private Sub pvHandleKey(KeyCode As Integer, Shift As Integer, ByVal bDown As Boolean)
    #If Shift Then
    #End If
    Select Case KeyCode
    Case vbKeyEscape
        m_isRunning = False
    Case vbKeyW, vbKeyUp
        m_keyIsDown(GameActionMoveUp) = bDown
    Case vbKeyA, vbKeyLeft
        m_keyIsDown(GameActionMoveLeft) = bDown
    Case vbKeyS, vbKeyDown
        m_keyIsDown(GameActionMoveDown) = bDown
    Case vbKeyD, vbKeyRight
        m_keyIsDown(GameActionMoveRight) = bDown
    End Select
End Sub

Private Sub Form_Load()
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
    Call IIDFromString(szIID_IDXGIFactory2, aGUID(0))
    Set dxgiFactory = dxgiAdapter.GetParent(aGUID(0))
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
    Call IIDFromString(szIID_ID3D11Texture2D, aGUID(0))
    Set d3d11FrameBuffer = m_d3d11SwapChain.GetBuffer(0, aGUID(0))
    Set m_d3d11FrameBufferView = m_d3d11Device.CreateRenderTargetView(d3d11FrameBuffer, ByVal 0)
    Set d3d11FrameBuffer = Nothing
    
    '--- Create Vertex Shader
    Dim shaderCompileErrorsBlob As ID3DBlob
    Dim errorString     As String
    hResult = D3DCompileFromFile(PathCombine(App.Path, "shaders.hlsl"), ByVal 0, ByVal 0, "vs_main", "vs_5_0", 0, 0, m_vsBlob, shaderCompileErrorsBlob)
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
    hResult = D3DCompileFromFile(PathCombine(App.Path, "shaders.hlsl"), ByVal 0, ByVal 0, "ps_main", "ps_5_0", 0, 0, psBlob, shaderCompileErrorsBlob)
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
    Dim nameBuffer(0 To 0) As UcsBuffer
    pvInitInputElementDesc inputElementDesc(0), nameBuffer(0), "POS", 0, DXGI_FORMAT_R32G32_FLOAT, 0, 0, D3D11_INPUT_PER_VERTEX_DATA, 0
    Set m_inputLayout = m_d3d11Device.CreateInputLayout(inputElementDesc(0), UBound(inputElementDesc) + 1, m_vsBlob.GetBufferPointer(), m_vsBlob.GetBufferSize())
    
    '--- Create Vertex Buffer
    Dim vertexData() As Single '--- x, y
    pvArraySingle vertexData, _
        -0.5!, 0.5!, _
        0.5!, -0.5!, _
        -0.5!, -0.5!, _
        -0.5!, 0.5!, _
        0.5!, 0.5!, _
        0.5!, -0.5!
    m_stride = 2 * 4
    m_numVerts = (UBound(vertexData) + 1) / 2
    m_offset = 0
    Dim vertexBufferDesc As D3D11_BUFFER_DESC
    With vertexBufferDesc
        .ByteWidth = (UBound(vertexData) + 1) * 4
        .Usage = D3D11_USAGE_IMMUTABLE
        .BindFlags = D3D11_BIND_VERTEX_BUFFER
    End With
    Dim vertexSubresourceData As D3D11_SUBRESOURCE_DATA
    vertexSubresourceData.pSysMem = VarPtr(vertexData(0))
    Set m_vertexBuffer = m_d3d11Device.CreateBuffer(vertexBufferDesc, vertexSubresourceData)
    
    '--- Create Constant Buffer
    Dim constantBufferDesc As D3D11_BUFFER_DESC
    With constantBufferDesc
        .ByteWidth = (sizeof_UcsConstants + &HF) And &HFFFFFFF0 '--- ByteWidth must be a multiple of 16, per the docs
        .Usage = D3D11_USAGE_DYNAMIC
        .BindFlags = D3D11_BIND_CONSTANT_BUFFER
        .CPUAccessFlags = D3D11_CPU_ACCESS_WRITE
    End With
    Set m_constantBuffer = m_d3d11Device.CreateBuffer(constantBufferDesc, ByVal 0)
    
    '--- Timing
    Dim currentTimeInSeconds As Double
    
    '--- Main Loop
    Show
    m_isRunning = True
    Do While m_isRunning
        Dim dt              As Single
        Dim previousTimeInSeconds As Double
        previousTimeInSeconds = currentTimeInSeconds
        currentTimeInSeconds = TimerEx
        dt = currentTimeInSeconds - previousTimeInSeconds
        If dt > 1! / 60! Then
            dt = 1! / 60!
        End If
        
        If m_windowDidResize Then
            m_d3d11DeviceContext.OMSetRenderTargets 0, Nothing, Nothing
            Set m_d3d11FrameBufferView = Nothing
            m_d3d11SwapChain.ResizeBuffers 0, 0, 0, DXGI_FORMAT_UNKNOWN, 0
            Call IIDFromString(szIID_ID3D11Texture2D, aGUID(0))
            Set d3d11FrameBuffer = m_d3d11SwapChain.GetBuffer(0, aGUID(0))
            Set m_d3d11FrameBufferView = m_d3d11Device.CreateRenderTargetView(d3d11FrameBuffer, ByVal 0)
            Set d3d11FrameBuffer = Nothing
            m_windowDidResize = False
        End If
        
        '--- Player movement logic
        Dim playerSpeed     As Single
        Dim playerMoveAmount As Single
        playerSpeed = 1.5!
        playerMoveAmount = playerSpeed * dt
        If m_keyIsDown(GameActionMoveUp) Then
            m_playerPos.y = m_playerPos.y + playerMoveAmount
        End If
        If m_keyIsDown(GameActionMoveDown) Then
            m_playerPos.y = m_playerPos.y - playerMoveAmount
        End If
        If m_keyIsDown(GameActionMoveLeft) Then
            m_playerPos.x = m_playerPos.x - playerMoveAmount
        End If
        If m_keyIsDown(GameActionMoveRight) Then
            m_playerPos.x = m_playerPos.x + playerMoveAmount
        End If
        
        '--- Cycle player color
        Dim playerColor     As D3D_FLOAT4
        Const colorCyclePeriod As Single = 5!  '--- in seconds
        Const colorCycleFreq As Single = 2! * 3.141592! / colorCyclePeriod
        playerColor.x = 0.5! * (Sin(colorCycleFreq * currentTimeInSeconds) + 1!)
        playerColor.y = 1! - playerColor.x
        playerColor.z = 0!
        playerColor.w = 1!
        
        '--- Update constant buffer
        Dim mappedSubresource As D3D11_MAPPED_SUBRESOURCE
        m_d3d11DeviceContext.Map m_constantBuffer, 0, D3D11_MAP_WRITE_DISCARD, 0, mappedSubresource
        Dim constants As UcsConstants
        constants.pos = m_playerPos
        constants.color = playerColor
        Call CopyMemory(ByVal mappedSubresource.pData, constants, sizeof_UcsConstants)
        m_d3d11DeviceContext.Unmap m_constantBuffer, 0
        
        Dim backgroundColor() As Single
        pvArraySingle backgroundColor, 0.1!, 0.2!, 0.6!, 1!
        m_d3d11DeviceContext.ClearRenderTargetView m_d3d11FrameBufferView, backgroundColor(0)
        
        Dim winRect As D3D11_RECT
        Call GetClientRect(hWnd, winRect)
        Dim viewport As D3D11_VIEWPORT
        With viewport
            .Width = winRect.Right - winRect.Left
            .Height = winRect.Bottom - winRect.Top
            .MaxDepth = 1
        End With
        m_d3d11DeviceContext.RSSetViewports 1, viewport
        
        m_d3d11DeviceContext.OMSetRenderTargets 1, m_d3d11FrameBufferView, Nothing
        
        m_d3d11DeviceContext.IASetPrimitiveTopology D3D11_PRIMITIVE_TOPOLOGY_TRIANGLELIST
        m_d3d11DeviceContext.IASetInputLayout m_inputLayout
        
        m_d3d11DeviceContext.VSSetShader m_vertexShader, Nothing, 0
        m_d3d11DeviceContext.PSSetShader m_pixelShader, Nothing, 0
        
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

Private Sub pvArrayLong(aDest() As Long, ParamArray A() As Variant)
    Dim lIdx            As Long
    
    ReDim aDest(0 To UBound(A)) As Long
    For lIdx = 0 To UBound(A)
        aDest(lIdx) = A(lIdx)
    Next
End Sub

Private Sub pvArraySingle(aDest() As Single, ParamArray A() As Variant)
    Dim lIdx            As Long
    
    ReDim aDest(0 To UBound(A)) As Single
    For lIdx = 0 To UBound(A)
        aDest(lIdx) = A(lIdx)
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
