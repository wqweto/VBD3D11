Attribute VB_Name = "md3DMaths"
Option Explicit
DefObj A-Z

Public Const M_PI As Double = 3.14159265

Public Function DegreesToRadians(ByVal sngDegs As Single) As Single
    DegreesToRadians = sngDegs * (M_PI / 180!)
End Function

Public Function XmMake3(ByVal x As Single, ByVal y As Single, ByVal z As Single) As XMFLOAT3
    With XmMake3
        .x = x
        .y = y
        .z = z
    End With
End Function

Public Function XmLength(uV As XMFLOAT3) As Single
    With uV
        XmLength = Sqr(.x * .x + .y * .y + .z * .z)
    End With
End Function

Public Function XmDot(uA As XMFLOAT4, uB As XMFLOAT4) As Single
    With uA
        XmDot = .x * uB.x + .y * uB.y + .z * uB.z + .w * uB.w
    End With
End Function

Public Function XmScalarMul(uV As XMFLOAT3, ByVal sngF As Single) As XMFLOAT3
    With uV
        XmScalarMul.x = .x * sngF
        XmScalarMul.y = .y * sngF
        XmScalarMul.z = .z * sngF
    End With
End Function

Public Function XmNormalize(uV As XMFLOAT3) As XMFLOAT3
    XmNormalize = XmScalarMul(uV, 1! / XmLength(uV))
End Function

Public Function XmCross(uA As XMFLOAT3, uB As XMFLOAT3) As XMFLOAT3
    With uA
        XmCross.x = .y * uB.z - .z * uB.y
        XmCross.y = .z * uB.x - .x * uB.z
        XmCross.z = .x * uB.y - .y * uB.x
    End With
End Function

Public Function XmAdd(uA As XMFLOAT3, uB As XMFLOAT3) As XMFLOAT3
    With uA
        XmAdd.x = .x + uB.x
        XmAdd.y = .y + uB.y
        XmAdd.z = .z + uB.z
    End With
End Function

Public Function XmSub(uA As XMFLOAT3, uB As XMFLOAT3) As XMFLOAT3
    With uA
        XmSub.x = .x - uB.x
        XmSub.y = .y - uB.y
        XmSub.z = .z - uB.z
    End With
End Function

Public Function XmNeg(uA As XMFLOAT3) As XMFLOAT3
    With uA
        XmNeg.x = -.x
        XmNeg.y = -.y
        XmNeg.z = -.z
    End With
End Function

Public Function XmRotateXMat(ByVal sngRad As Single) As XMMATRIX
    Dim sngSin          As Single
    Dim sngCos          As Single
    
    sngSin = Sin(sngRad)
    sngCos = Cos(sngRad)
    With XmRotateXMat
        .m(0, 0) = 1
                            .m(1, 1) = sngCos:  .m(2, 1) = -sngSin
                            .m(1, 2) = sngSin:  .m(2, 2) = sngCos
                                                                    .m(3, 3) = 1
    End With
End Function

Public Function XmRotateYMat(ByVal sngRad As Single) As XMMATRIX
    Dim sngSin          As Single
    Dim sngCos          As Single
    
    sngSin = Sin(sngRad)
    sngCos = Cos(sngRad)
    With XmRotateYMat
        .m(0, 0) = sngCos:                      .m(2, 0) = sngSin
                            .m(1, 1) = 1
        .m(0, 2) = -sngSin:                     .m(2, 2) = sngCos
                                                                    .m(3, 3) = 1
    End With
End Function

Public Function XmTranslationMat(uTrans As XMFLOAT3) As XMMATRIX
    With XmTranslationMat
        .m(0, 0) = 1:                                               .m(3, 0) = uTrans.x
                            .m(1, 1) = 1:                           .m(3, 1) = uTrans.y
                                                .m(2, 2) = 1:       .m(3, 2) = uTrans.z
                                                                    .m(3, 3) = 1
    End With
End Function

Public Function XmMakePerspectiveMat(ByVal sngAspectRatio As Single, ByVal sngFovYRadians As Single, ByVal sngZNear As Single, ByVal sngZFar As Single) As XMMATRIX
    Dim yScale          As Single
    Dim xScale          As Single
    Dim zRangeInverse   As Single
    Dim zScale          As Single
    Dim zTranslation    As Single
    
    yScale = Tan(0.5! * (M_PI - sngFovYRadians))
    xScale = yScale / sngAspectRatio
    zRangeInverse = 1! / (sngZNear - sngZFar)
    zScale = sngZFar * zRangeInverse
    zTranslation = sngZFar * sngZNear * zRangeInverse
    With XmMakePerspectiveMat
        .m(0, 0) = xScale:
                            .m(1, 1) = yScale:
                                                .m(2, 2) = zScale:  .m(3, 2) = zTranslation
                                                .m(2, 3) = -1
    End With
End Function

Public Function XmColMat(uMat As XMMATRIX, ByVal Index As Long) As XMFLOAT4
    With uMat
        XmColMat.x = .m(0, Index)
        XmColMat.y = .m(1, Index)
        XmColMat.z = .m(2, Index)
        XmColMat.w = .m(3, Index)
    End With
End Function

Public Function XmRowMat(uMat As XMMATRIX, ByVal Index As Long) As XMFLOAT4
    With uMat
        XmRowMat.x = .m(Index, 0)
        XmRowMat.y = .m(Index, 1)
        XmRowMat.z = .m(Index, 2)
        XmRowMat.w = .m(Index, 3)
    End With
End Function

Public Function XmMulMat(uA As XMMATRIX, uB As XMMATRIX) As XMMATRIX
    Dim uRow0           As XMFLOAT4
    Dim uRow1           As XMFLOAT4
    Dim uRow2           As XMFLOAT4
    Dim uRow3           As XMFLOAT4
    Dim uCol            As XMFLOAT4
    
    uRow0 = XmRowMat(uA, 0)
    uRow1 = XmRowMat(uA, 1)
    uRow2 = XmRowMat(uA, 2)
    uRow3 = XmRowMat(uA, 3)
    With XmMulMat
        uCol = XmColMat(uB, 0)
        .m(0, 0) = XmDot(uRow0, uCol)
        .m(1, 0) = XmDot(uRow1, uCol)
        .m(2, 0) = XmDot(uRow2, uCol)
        .m(3, 0) = XmDot(uRow3, uCol)
        uCol = XmColMat(uB, 1)
        .m(0, 1) = XmDot(uRow0, uCol)
        .m(1, 1) = XmDot(uRow1, uCol)
        .m(2, 1) = XmDot(uRow2, uCol)
        .m(3, 1) = XmDot(uRow3, uCol)
        uCol = XmColMat(uB, 2)
        .m(0, 2) = XmDot(uRow0, uCol)
        .m(1, 2) = XmDot(uRow1, uCol)
        .m(2, 2) = XmDot(uRow2, uCol)
        .m(3, 2) = XmDot(uRow3, uCol)
        uCol = XmColMat(uB, 3)
        .m(0, 3) = XmDot(uRow0, uCol)
        .m(1, 3) = XmDot(uRow1, uCol)
        .m(2, 3) = XmDot(uRow2, uCol)
        .m(3, 3) = XmDot(uRow3, uCol)
    End With
End Function
