Attribute VB_Name = "RayCastingAlgo"


Option Explicit

Function RayCastingAlgo3D(TestPoint As Variant, vPlanarFaceName As Variant, _
                        FaceNames_Fd As Range, EdgeTypes_Fd As Range, Vertex1 As Variant, Vertex2 As Variant, _
                        FaceNames_FV As Range, FaceNormals_FV As Variant, EdgeCurvatures_Fd As Range) As Variant

    
    If ((TestPoint(1, 1) = "" And TestPoint(1, 2) = "" And TestPoint(1, 3) = "") Or vPlanarFaceName = "") Then
        RayCastingAlgo3D = ""
    End If
    
    Dim vBoundaryPt As Boolean
    vBoundaryPt = fnIsBoundaryPoint(TestPoint, vPlanarFaceName, FaceNames_Fd, EdgeTypes_Fd, Vertex1, Vertex2, EdgeCurvatures_Fd)
    
    If (vBoundaryPt = True) Then
        RayCastingAlgo3D = True
        Exit Function
    End If
 
    Dim i As Integer, j As Integer
    Dim sRow As Integer, eRow As Integer
    sRow = GetAttrib_StartRow(vPlanarFaceName, FaceNames_Fd)
    eRow = GetAttrib_EndRow(vPlanarFaceName, FaceNames_Fd)
    
    Dim LinEdgeVtx1 As Variant, LinEdgeVtx2 As Variant, LinEdgeVector As Variant, LinEdgeDCs As Variant
    ReDim LinEdgeVtx1(1 To 1, 1 To 3), LinEdgeVtx2(1 To 1, 1 To 3), LinEdgeVector(1 To 1, 1 To 3), LinEdgeDCs(1 To 1, 1 To 3)
    Dim sEdgeType As String
    For i = sRow To eRow
        sEdgeType = UCase(EdgeTypes_Fd(i))
        If (sEdgeType = "LINEAR" Or sEdgeType = "INTERSECTION" Or sEdgeType = "SPCURVE") Then
            For j = 1 To 3
                LinEdgeVtx1(1, j) = Vertex1(i, j)
                LinEdgeVtx2(1, j) = Vertex2(i, j)
                LinEdgeVector(1, j) = LinEdgeVtx2(1, j) - LinEdgeVtx1(1, j)
            Next
            LinEdgeDCs = GetDCs(LinEdgeVector(1, 1), LinEdgeVector(1, 2), LinEdgeVector(1, 3))
            Exit For
        End If
    Next

    Dim lRow_Face As Long
    Dim CurrFaceNormal As Variant
    ReDim CurrFaceNormal(1 To 1, 1 To 3)
    lRow_Face = WorksheetFunction.IfError(Application.Match(vPlanarFaceName, FaceNames_FV, 0), 0)
    If (lRow_Face > 0) Then
        For i = 1 To 3
            CurrFaceNormal(1, i) = FaceNormals_FV(lRow_Face, i)
        Next
    End If

    Dim PerpDirCosines As Variant
    ReDim PerpDirCosines(1 To 1, 1 To 3)
    PerpDirCosines = fnVectorCrossProduct(CurrFaceNormal, LinEdgeDCs)
    
    Dim vEpsilonDist As Double
    vEpsilonDist = dCoordTol
    Dim vTolerance As Double
    vTolerance = dCoordTol
    
    Dim EdgeVtx1 As Variant, EdgeVtx2 As Variant, IntersectionPt As Variant
    ReDim EdgeVtx1(1 To 1, 1 To 3), EdgeVtx2(1 To 1, 1 To 3), IntersectionPt(1 To 1, 1 To 3)
    Dim vDist1 As Double, vDist2 As Double, MaxDist As Double
    MaxDist = vInitMax

    For i = sRow To eRow
        For j = 1 To 3
            EdgeVtx1(1, j) = Vertex1(i, j)
            EdgeVtx2(1, j) = Vertex2(i, j)
        Next
        
        If (Not ((Abs(EdgeVtx1(1, 1) - EdgeVtx2(1, 1)) < vTolerance) And _
                 (Abs(EdgeVtx1(1, 2) - EdgeVtx2(1, 2)) < vTolerance) And _
                 (Abs(EdgeVtx1(1, 3) - EdgeVtx2(1, 3)) < vTolerance))) Then
            
            vDist1 = ((EdgeVtx1(1, 1) - TestPoint(1, 1)) * LinEdgeDCs(1, 1)) + _
                     ((EdgeVtx1(1, 2) - TestPoint(1, 2)) * LinEdgeDCs(1, 2)) + _
                     ((EdgeVtx1(1, 3) - TestPoint(1, 3)) * LinEdgeDCs(1, 3))
            vDist2 = ((EdgeVtx2(1, 1) - TestPoint(1, 1)) * LinEdgeDCs(1, 1)) + _
                     ((EdgeVtx2(1, 2) - TestPoint(1, 2)) * LinEdgeDCs(1, 2)) + _
                     ((EdgeVtx2(1, 3) - TestPoint(1, 3)) * LinEdgeDCs(1, 3))
            If (vDist1 >= MaxDist) Then
                MaxDist = vDist1
            End If
            If (vDist2 >= MaxDist) Then
                MaxDist = vDist2
            End If
        End If
    Next
   
    Dim vOffset  As Double
    vOffset = 100
    Dim RayStPt As Variant, RayEndPt As Variant
    ReDim RayStPt(1 To 1, 1 To 3), RayEndPt(1 To 1, 1 To 3)
    For i = 1 To 3
        RayStPt(1, i) = TestPoint(1, i)
        RayEndPt(1, i) = TestPoint(1, i) + ((Abs(MaxDist) + vOffset) * LinEdgeDCs(1, i))
    Next
 
    Dim IntersectionsCollection As New Collection
    Dim vIntersectionCount As Integer
    
    Dim sEdgeCurvature As String
    
    For i = sRow To eRow
        
        For j = 1 To 3
            EdgeVtx1(1, j) = Vertex1(i, j)
            EdgeVtx2(1, j) = Vertex2(i, j)
        Next

        If (Not ((Abs(EdgeVtx1(1, 1) - EdgeVtx2(1, 1)) < vTolerance) And _
                 (Abs(EdgeVtx1(1, 2) - EdgeVtx2(1, 2)) < vTolerance) And _
                 (Abs(EdgeVtx1(1, 3) - EdgeVtx2(1, 3)) < vTolerance))) Then
            
            sEdgeType = EdgeTypes_Fd(i)
            sEdgeCurvature = EdgeCurvatures_Fd(i)
            If (UCase(sEdgeType) = "CIRCULAR" And UCase(sEdgeCurvature) = "CONVEX") Then
               
                 Such line vectors do not form Body's outline & hence not relevant for Intersctions calculation. _
                 Whereas, Concave Circular edges contribute to Body's outline.
            
            Else

                IntersectionPt = fnComputeIntersectionPt_TwoLineSegments(RayStPt, RayEndPt, EdgeVtx1, EdgeVtx2)
                
                If (IntersectionPt(1, 1) <> "" And IntersectionPt(1, 2) <> "" And IntersectionPt(1, 3) <> "") Then
                    
                    If ((Abs(IntersectionPt(1, 1) - EdgeVtx1(1, 1)) < vTolerance) And _
                        (Abs(IntersectionPt(1, 2) - EdgeVtx1(1, 2)) < vTolerance) And _
                        (Abs(IntersectionPt(1, 3) - EdgeVtx1(1, 3)) < vTolerance)) Then

                        For j = 1 To 3
                            EdgeVtx1(1, j) = EdgeVtx1(1, j) + (vEpsilonDist * PerpDirCosines(1, j))
                        Next
                        IntersectionPt = fnComputeIntersectionPt_TwoLineSegments(RayStPt, RayEndPt, EdgeVtx1, EdgeVtx2)
                    
                    ElseIf ((Abs(IntersectionPt(1, 1) - EdgeVtx2(1, 1)) < vTolerance) And _
                            (Abs(IntersectionPt(1, 2) - EdgeVtx2(1, 2)) < vTolerance) And _
                            (Abs(IntersectionPt(1, 3) - EdgeVtx2(1, 3)) < vTolerance)) Then
                    

                        For j = 1 To 3
                            EdgeVtx2(1, j) = EdgeVtx2(1, j) + (vEpsilonDist * PerpDirCosines(1, j))
                        Next
                        IntersectionPt = fnComputeIntersectionPt_TwoLineSegments(RayStPt, RayEndPt, EdgeVtx1, EdgeVtx2)
                        
                    End If
                    
                    If (IntersectionPt(1, 1) <> "" And IntersectionPt(1, 2) <> "" And IntersectionPt(1, 3) <> "") Then
                        IntersectionsCollection.Add IntersectionPt
                        vIntersectionCount = vIntersectionCount + 1
                    End If
                    
                End If
            End If
        End If
    Next
    
    If (vIntersectionCount Mod 2 = 0) Then
        RayCastingAlgo3D = False
    ElseIf (vIntersectionCount Mod 2 = 1) Then
        RayCastingAlgo3D = True
    End If

End Function

Function fnIsBoundaryPoint(TestPoint As Variant, vPlanarFaceName As Variant, FaceNames_Fd As Range, _
            EdgeTypes_Fd As Range, Vertex1 As Variant, Vertex2 As Variant, EdgeCurvatures_Fd As Range) As Variant

    Dim sRow As Integer, eRow As Integer
    sRow = GetAttrib_StartRow(vPlanarFaceName, FaceNames_Fd)
    eRow = GetAttrib_EndRow(vPlanarFaceName, FaceNames_Fd)
    
    Dim i As Integer, j As Integer
    Dim FirstPt As Variant, SecondPt As Variant, ThirdPt As Variant
    ReDim FirstPt(1 To 1, 1 To 3), SecondPt(1 To 1, 1 To 3), ThirdPt(1 To 1, 1 To 3)
    
    Dim vTolerance As Double, vUnitVecTol As Double
    vTolerance = dCoordTol

    vUnitVecTol = dAngTol_1deg_Prl
    
    
    Dim aVector As Variant, bVector As Variant
    ReDim aVector(1 To 1, 1 To 3), bVector(1 To 1, 1 To 3)
    Dim aUnitVector As Variant, bUnitVector As Variant
    ReDim aUnitVector(1 To 1, 1 To 3), bUnitVector(1 To 1, 1 To 3)
    
    Dim sEdgeType As String
    Dim sEdgeCurvature As String
    
    For i = sRow To eRow
        
        sEdgeType = EdgeTypes_Fd(i)
        sEdgeCurvature = EdgeCurvatures_Fd(i)
        If (UCase(sEdgeType) = "CIRCULAR" And UCase(sEdgeCurvature) = "CONVEX") Then

             Such line vectors do not form Body's outline & hence not relevant for Intersctions calculation. _
             Whereas, Concave Circular edges contribute to Body's outline.
            
            
        Else
            For j = 1 To 3
                FirstPt(1, j) = Vertex1(i, j)
                SecondPt(1, j) = TestPoint(1, j)
                ThirdPt(1, j) = Vertex2(i, j)
                aVector(1, j) = FirstPt(1, j) - SecondPt(1, j)
                bVector(1, j) = ThirdPt(1, j) - SecondPt(1, j)
            Next

            If (((Abs(aVector(1, 1)) < vTolerance) And (Abs(aVector(1, 2)) < vTolerance) And (Abs(aVector(1, 3)) < vTolerance)) Or _
                ((Abs(bVector(1, 1)) < vTolerance) And (Abs(bVector(1, 2)) < vTolerance) And (Abs(bVector(1, 3)) < vTolerance))) Then
                fnIsBoundaryPoint = True
                Exit Function
            End If
            
            aUnitVector = GetDCs(aVector(1, 1), aVector(1, 2), aVector(1, 3))
            bUnitVector = GetDCs(bVector(1, 1), bVector(1, 2), bVector(1, 3))
 
            If ((Abs(aUnitVector(1, 1) + bUnitVector(1, 1)) < vUnitVecTol) And _
                (Abs(aUnitVector(1, 2) + bUnitVector(1, 2)) < vUnitVecTol) And _
                (Abs(aUnitVector(1, 3) + bUnitVector(1, 3)) < vUnitVecTol)) Then
                fnIsBoundaryPoint = True
                Exit Function
            End If
        End If
    Next
    
    fnIsBoundaryPoint = False
    
End Function

Function fnComputeIntersectionPt_TwoLineSegments(E1_Vertex1 As Variant, E1_Vertex2 As Variant, _
                                                    E2_Vertex1 As Variant, E2_Vertex2 As Variant) As Variant


    Dim i As Integer
    Dim IntersectionPt As Variant
    ReDim IntersectionPt(1 To 1, 1 To 3)
    For i = 1 To 3
        IntersectionPt(1, i) = ""
    Next

    Dim EVector1 As Variant, EVector2 As Variant
    ReDim EVector1(1 To 1, 1 To 3), EVector2(1 To 1, 1 To 3)
    For i = 1 To 3
        EVector1(1, i) = E1_Vertex2(1, i) - E1_Vertex1(1, i)
        EVector2(1, i) = E2_Vertex2(1, i) - E2_Vertex1(1, i)
    Next

    Dim CrossProComps As Variant
    ReDim CrossProComps(1 To 1, 1 To 3)
    CrossProComps = fnVectorCrossProduct(EVector1, EVector2)

    Dim vTolerance As Double
    vTolerance = 0.00001

    If ((Abs(CrossProComps(1, 1)) < vTolerance) And (Abs(CrossProComps(1, 2)) < vTolerance) And (Abs(CrossProComps(1, 3)) < vTolerance)) Then
        fnComputeIntersectionPt_TwoLineSegments = IntersectionPt
        Exit Function
    End If
    Dim vPlaneNormalDCs As Variant
    ReDim vPlaneNormalDCs(1 To 1, 1 To 3)

    vPlaneNormalDCs = fnVectorCrossProduct(EVector1, EVector2)

    Dim LinCombVector As Variant
    ReDim LinCombVector(1 To 1, 1 To 3)
    LinCombVector = fnVectorCrossProduct(vPlaneNormalDCs, EVector2)

    
    
    Dim vMagnLinCombVec As Double, vMagnEdge2 As Double
    vMagnLinCombVec = fnVectorLength(LinCombVector(1, 1), LinCombVector(1, 2), LinCombVector(1, 3))
    vMagnEdge2 = fnVectorLength(EVector2(1, 1), EVector2(1, 2), EVector2(1, 3))

    Dim LeftNormalVec As Variant
    ReDim LeftNormalVec(1 To 1, 1 To 3)
    For i = 1 To 3
        LeftNormalVec(1, i) = LinCombVector(1, i) * (vMagnEdge2 / vMagnLinCombVec)
    Next

    Dim TempVect As Variant
    ReDim TempVect(1 To 1, 1 To 3)
    For i = 1 To 3
        TempVect(1, i) = E2_Vertex1(1, i) - E1_Vertex1(1, i)
    Next

    Dim vDotPro1 As Double, vDotPro2 As Double
    vDotPro1 = (TempVect(1, 1) * LeftNormalVec(1, 1)) + (TempVect(1, 2) * LeftNormalVec(1, 2)) + (TempVect(1, 3) * LeftNormalVec(1, 3))
    vDotPro2 = (EVector1(1, 1) * LeftNormalVec(1, 1)) + (EVector1(1, 2) * LeftNormalVec(1, 2)) + (EVector1(1, 3) * LeftNormalVec(1, 3))

    Dim vScalarT As Double
    vScalarT = vDotPro1 / vDotPro2

    LinCombVector = fnVectorCrossProduct(vPlaneNormalDCs, EVector1)


    Dim vMagnEdge1 As Double
    vMagnLinCombVec = fnVectorLength(LinCombVector(1, 1), LinCombVector(1, 2), LinCombVector(1, 3))
    vMagnEdge1 = fnVectorLength(EVector1(1, 1), EVector1(1, 2), EVector1(1, 3))

    For i = 1 To 3
        LeftNormalVec(1, i) = LinCombVector(1, i) * (vMagnEdge1 / vMagnLinCombVec)
    Next
    For i = 1 To 3
        TempVect(1, i) = E1_Vertex1(1, i) - E2_Vertex1(1, i)
    Next
    vDotPro1 = (TempVect(1, 1) * LeftNormalVec(1, 1)) + (TempVect(1, 2) * LeftNormalVec(1, 2)) + (TempVect(1, 3) * LeftNormalVec(1, 3))
    vDotPro2 = (EVector2(1, 1) * LeftNormalVec(1, 1)) + (EVector2(1, 2) * LeftNormalVec(1, 2)) + (EVector2(1, 3) * LeftNormalVec(1, 3))

    Dim vScalarS As Double
    vScalarS = vDotPro1 / vDotPro2

    If ((vScalarT >= 0 And vScalarT <= 1) And _
        (vScalarS >= 0 And vScalarS <= 1)) Then
        For i = 1 To 3
            IntersectionPt(1, i) = E1_Vertex1(1, i) + (vScalarT * EVector1(1, i))
        Next
    End If

    fnComputeIntersectionPt_TwoLineSegments = IntersectionPt

End Function





