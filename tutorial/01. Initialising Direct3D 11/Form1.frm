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

Private m_d3d11Device           As ID3D11Device1
Private m_d3d11DeviceContext    As ID3D11DeviceContext1
Private m_d3d11SwapChain        As IDXGISwapChain1
Private m_d3d11FrameBufferView  As ID3D11RenderTargetView
Private m_isRunning             As Boolean

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
    
    '--- Main Loop
    Show
    m_isRunning = True
    Do While m_isRunning
        DoEvents
        Dim backgroundColor() As Single
        pvArraySingle backgroundColor, 0.1!, 0.2!, 0.6!, 1!
        m_d3d11DeviceContext.ClearRenderTargetView m_d3d11FrameBufferView, backgroundColor(0)
        m_d3d11SwapChain.Present 1, 0
    Loop
End Sub

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_isRunning = False
End Sub
