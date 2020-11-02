Attribute VB_Name = "mdObjLoading"
Option Explicit

Private Const EPSILON           As Single = 0.0001

Public Type UcsVertexDataType
    pos(0 To 3)         As Single
    uv(0 To 1)          As Single
    norm(0 To 3)        As Single
End Type

Public Type UcsLoadedObjType
    numVertexes         As Long
    numIndices          As Long
    
    vertexBuffer()      As UcsVertexDataType
    indexBuffer()       As Integer
End Type

Public Function LoadObj(sFileName As String) As UcsLoadedObjType
    Dim vLines          As Variant
    Dim lIdx            As Long
    Dim vSplit          As Variant
    Dim numPositions    As Long
    Dim numTexCoords    As Long
    Dim numNormals      As Long
    Dim cVertex         As Collection
    Dim lJdx            As Long
    Dim vFace           As Variant
    Dim uData           As UcsVertexDataType
    Dim lIndex          As Long
    Dim bSmoothNormals  As Boolean
    Dim sngInvLength    As Single
    
    With LoadObj
        ReDim .vertexBuffer(0 To 32) As UcsVertexDataType
        ReDim .indexBuffer(0 To 32) As Integer
        Set cVertex = New Collection
        vLines = Split(ReadTextFile(sFileName), vbCrLf)
        For lIdx = 0 To UBound(vLines)
            vSplit = Split(vLines(lIdx))
            Select Case LCase$(At(vSplit, 0))
            Case "v"
                numPositions = numPositions + 1
                cVertex.Add pvParseArray(vSplit, 3), "v" & numPositions
            Case "vt"
                numTexCoords = numTexCoords + 1
                cVertex.Add pvParseArray(vSplit, 2), "vt" & numTexCoords
            Case "vn"
                numNormals = numNormals + 1
                cVertex.Add pvParseArray(vSplit, 3), "vn" & numNormals
            Case "f"
                For lJdx = 1 To UBound(vSplit)
                    vFace = Split(vSplit(lJdx), "/")
                    pvAssignData uData, cVertex.Item("v" & At(vFace, 0)), cVertex.Item("vt" & At(vFace, 1)), cVertex.Item("vn" & At(vFace, 2))
                    lIndex = pvFindVertex(.vertexBuffer, .numVertexes, bSmoothNormals, uData)
                    If lIndex >= .numVertexes Then
                        If .numVertexes > UBound(.vertexBuffer) Then
                            ReDim Preserve .vertexBuffer(0 To UBound(.vertexBuffer) * 2) As UcsVertexDataType
                        End If
                        .vertexBuffer(.numVertexes) = uData
                        .numVertexes = .numVertexes + 1
                    ElseIf bSmoothNormals Then
                        With .vertexBuffer(lIndex)
                            .norm(0) = .norm(0) + uData.norm(0)
                            .norm(1) = .norm(1) + uData.norm(1)
                            .norm(2) = .norm(2) + uData.norm(2)
                        End With
                    End If
                    If .numIndices > UBound(.indexBuffer) Then
                        ReDim Preserve .indexBuffer(0 To UBound(.indexBuffer) * 2) As Integer
                    End If
                    .indexBuffer(.numIndices) = lIndex
                    .numIndices = .numIndices + 1
                Next
            Case "s"
                bSmoothNormals = LCase$(At(vSplit, 1)) = "on" Or Val(At(vSplit, 1)) <> 0
            End Select
        Next
        '--- Normalise the normals
        For lIdx = 0 To .numVertexes - 1
            With .vertexBuffer(lIdx)
                sngInvLength = Sqr(.norm(0) * .norm(0) + .norm(1) * .norm(1) + .norm(2) * .norm(2))
                If Abs(sngInvLength) > EPSILON Then
                    sngInvLength = 1 / sngInvLength
                End If
                .norm(0) = .norm(0) * sngInvLength
                .norm(1) = .norm(1) * sngInvLength
                .norm(2) = .norm(2) * sngInvLength
            End With
        Next
    End With
End Function

Private Function pvParseArray(vLine As Variant, ByVal NumEntries As Long) As Single()
    Dim aRetVal()       As Single
    Dim lIdx            As Long
    
    ReDim aRetVal(0 To NumEntries - 1) As Single
    For lIdx = 0 To NumEntries - 1
        aRetVal(lIdx) = Val(At(vLine, lIdx + 1))
    Next
    pvParseArray = aRetVal
End Function

Private Sub pvAssignData(uData As UcsVertexDataType, vPos As Variant, vTexCoords As Variant, vNorm As Variant)
    With uData
        .pos(0) = vPos(0)
        .pos(1) = vPos(1)
        .pos(2) = vPos(2)
        .uv(0) = vTexCoords(0)
        .uv(1) = vTexCoords(1)
        .norm(0) = vNorm(0)
        .norm(1) = vNorm(1)
        .norm(2) = vNorm(2)
    End With
End Sub

Private Function pvFindVertex(vertexBuffer() As UcsVertexDataType, ByVal numVertexes As Long, ByVal bSmoothNormals As Boolean, uData As UcsVertexDataType) As Long
    Dim lIdx            As Long
    
    For lIdx = 0 To numVertexes - 1
        With vertexBuffer(lIdx)
            If Abs(.pos(0) - uData.pos(0)) < EPSILON And Abs(.pos(1) - uData.pos(1)) < EPSILON And Abs(.pos(2) - uData.pos(2)) < EPSILON Then
                If Abs(.uv(0) - uData.uv(0)) < EPSILON And Abs(.uv(1) - uData.uv(1)) < EPSILON Then
                    If bSmoothNormals Then
                        Exit For
                    ElseIf Abs(.norm(0) - uData.norm(0)) < EPSILON And Abs(.norm(1) - uData.norm(1)) < EPSILON And Abs(.norm(2) - uData.norm(2)) < EPSILON Then
                        Exit For
                    End If
                End If
            End If
        End With
    Next
    pvFindVertex = lIdx
End Function

Public Function ReadTextFile(sFileName As String) As String
    Dim sCharset            As String

    sCharset = "utf-8"
    With CreateObject("ADODB.Stream")
        .Open
        If LenB(sCharset) <> 0 Then
            .Charset = sCharset
        End If
        .LoadFromFile sFileName
        ReadTextFile = .ReadText()
    End With
End Function

Public Property Get At(vData As Variant, ByVal lIdx As Long, Optional sDefault As String) As String
    On Error GoTo QH
    At = sDefault
    If IsArray(vData) Then
        If lIdx < LBound(vData) Then
            '--- lIdx = -1 for last element
            lIdx = UBound(vData) + 1 + lIdx
        End If
        If LBound(vData) <= lIdx And lIdx <= UBound(vData) Then
            At = CStr(vData(lIdx))
        End If
    End If
QH:
End Property
