Attribute VB_Name = "mMaths"
Option Explicit

' Declare local variables ...

    Private P_dCalculatorMatrix As D3DMATRIX    ' Matrix used for calculations

' MIDENTITY: Resets a Matrix to Identity
Public Sub MIdentity(dMatrix As D3DMATRIX)
    dMatrix.v_11 = 1
    dMatrix.v_12 = 0
    dMatrix.v_13 = 0
    dMatrix.v_14 = 0
    dMatrix.v_21 = 0
    dMatrix.v_22 = 1
    dMatrix.v_23 = 0
    dMatrix.v_24 = 0
    dMatrix.v_31 = 0
    dMatrix.v_32 = 0
    dMatrix.v_33 = 1
    dMatrix.v_34 = 0
    dMatrix.v_41 = 0
    dMatrix.v_42 = 0
    dMatrix.v_43 = 0
    dMatrix.v_44 = 1
End Sub

' MTRANSLATE: Translates a matrix along the axis
Public Function MTranslate(dSource As D3DMATRIX, ByVal nValueX As Single, ByVal nValueY As Single, ByVal nValueZ As Single) As D3DMATRIX

    ' Reset to identity matrix
    MIdentity P_dCalculatorMatrix
    
    ' Add calculations
    With P_dCalculatorMatrix
        .v_41 = nValueX
        .v_42 = nValueY
        .v_43 = nValueZ
    End With

    ' Apply transformations
    dSource = MMultiply(P_dCalculatorMatrix, dSource)

    ' Return result
    MTranslate = dSource
    
End Function

' MSCALE: Scales a matrix along the axis
Public Function MScale(dSource As D3DMATRIX, ByVal nValueX As Single, ByVal nValueY As Single, ByVal nValueZ As Single) As D3DMATRIX

    ' Reset to identity matrix
    MIdentity P_dCalculatorMatrix
    
    ' Add calculations
    With P_dCalculatorMatrix
        .v_11 = nValueX
        .v_22 = nValueY
        .v_33 = nValueZ
    End With
    
    ' Apply transformations
    dSource = MMultiply(P_dCalculatorMatrix, dSource)
    
    ' Return result
    MScale = dSource
    
End Function

' MROTATE: Rotates a matrix around the axis
Public Function MRotate(dSource As D3DMATRIX, ByVal nValueX As Single, ByVal nValueY As Single, ByVal nValueZ As Single) As D3DMATRIX

    ' Setup local variables ...
        
        Dim L_nCos As Single        ' Holds cosine value of angle
        Dim L_nSin As Single        ' Holds sine value of angle
        
    ' Do rotation around X-Axis ...
        If nValueX <> 0 Then
        
            ' Reset caluculator to identity matrix
            MIdentity P_dCalculatorMatrix
                
            ' Get angle values
            L_nCos = Cos(nValueX * PIFactor)
            L_nSin = Sin(nValueX * PIFactor)
            
            ' Add transformations
            With P_dCalculatorMatrix
                .v_22 = L_nCos
                .v_33 = L_nCos
                .v_23 = -L_nSin
                .v_32 = L_nSin
            End With
            
            ' Apply transformations
            dSource = MMultiply(P_dCalculatorMatrix, dSource)
            
        End If
    
    ' Do rotation around Y-Axis ...
        If nValueY <> 0 Then
        
            ' Reset caluculator to identity matrix
            MIdentity P_dCalculatorMatrix
            
            ' Get angle values
            L_nCos = Cos(nValueY * PIFactor)
            L_nSin = Sin(nValueY * PIFactor)
            
            ' Add transformations
            With P_dCalculatorMatrix
                .v_11 = L_nCos
                .v_33 = L_nCos
                .v_13 = L_nSin
                .v_31 = -L_nSin
            End With
            
            ' Apply transformations
            dSource = MMultiply(P_dCalculatorMatrix, dSource)
            
        End If
        
    ' Do rotation around Z-Axis ...
        If nValueZ <> 0 Then
        
            ' Reset caluculator to identity matrix
            MIdentity P_dCalculatorMatrix
            
            ' Get angle values
            L_nCos = Cos(nValueZ * PIFactor)
            L_nSin = Sin(nValueZ * PIFactor)
            
            ' Add transformations
            With P_dCalculatorMatrix
                .v_11 = L_nCos
                .v_22 = L_nCos
                .v_12 = -L_nSin
                .v_21 = L_nSin
            End With
            
            ' Apply transformations
            dSource = MMultiply(P_dCalculatorMatrix, dSource)
            
        End If
        
    ' Return result ...
        MRotate = dSource
    
End Function

' MMULTIPLY: Multiplies two matrices
Public Function MMultiply(dM1 As D3DMATRIX, dM2 As D3DMATRIX) As D3DMATRIX

    ' Calculate multiply ...
        With MMultiply
            
            .v_11 = dM1.v_11 * dM2.v_11 + dM1.v_21 * dM2.v_12 + dM1.v_31 * dM2.v_13 + dM1.v_41 * dM2.v_14
            .v_21 = dM1.v_11 * dM2.v_21 + dM1.v_21 * dM2.v_22 + dM1.v_31 * dM2.v_23 + dM1.v_41 * dM2.v_24
            .v_31 = dM1.v_11 * dM2.v_31 + dM1.v_21 * dM2.v_32 + dM1.v_31 * dM2.v_33 + dM1.v_41 * dM2.v_34
            .v_41 = dM1.v_11 * dM2.v_41 + dM1.v_21 * dM2.v_42 + dM1.v_31 * dM2.v_43 + dM1.v_41 * dM2.v_44
    
            .v_12 = dM1.v_12 * dM2.v_11 + dM1.v_22 * dM2.v_12 + dM1.v_32 * dM2.v_13 + dM1.v_42 * dM2.v_14
            .v_22 = dM1.v_12 * dM2.v_21 + dM1.v_22 * dM2.v_22 + dM1.v_32 * dM2.v_23 + dM1.v_42 * dM2.v_24
            .v_32 = dM1.v_12 * dM2.v_31 + dM1.v_22 * dM2.v_32 + dM1.v_32 * dM2.v_33 + dM1.v_42 * dM2.v_34
            .v_42 = dM1.v_12 * dM2.v_41 + dM1.v_22 * dM2.v_42 + dM1.v_32 * dM2.v_43 + dM1.v_42 * dM2.v_44
    
            .v_13 = dM1.v_13 * dM2.v_11 + dM1.v_23 * dM2.v_12 + dM1.v_33 * dM2.v_13 + dM1.v_43 * dM2.v_14
            .v_23 = dM1.v_13 * dM2.v_21 + dM1.v_23 * dM2.v_22 + dM1.v_33 * dM2.v_23 + dM1.v_43 * dM2.v_24
            .v_33 = dM1.v_13 * dM2.v_31 + dM1.v_23 * dM2.v_32 + dM1.v_33 * dM2.v_33 + dM1.v_43 * dM2.v_34
            .v_43 = dM1.v_13 * dM2.v_41 + dM1.v_23 * dM2.v_42 + dM1.v_33 * dM2.v_43 + dM1.v_43 * dM2.v_44
    
            .v_14 = dM1.v_14 * dM2.v_11 + dM1.v_24 * dM2.v_12 + dM1.v_34 * dM2.v_13 + dM1.v_44 * dM2.v_14
            .v_24 = dM1.v_14 * dM2.v_21 + dM1.v_24 * dM2.v_22 + dM1.v_34 * dM2.v_23 + dM1.v_44 * dM2.v_24
            .v_34 = dM1.v_14 * dM2.v_31 + dM1.v_24 * dM2.v_32 + dM1.v_34 * dM2.v_33 + dM1.v_44 * dM2.v_34
            .v_44 = dM1.v_14 * dM2.v_41 + dM1.v_24 * dM2.v_42 + dM1.v_34 * dM2.v_43 + dM1.v_44 * dM2.v_44
    
        End With
            
    
End Function

' MDEBUG: Print a given matrix in the debug window
Public Sub MDebug(dMatrix As D3DMATRIX)

    With dMatrix
        Debug.Print Format(.v_11, "0.00000") + " " + Format(.v_12, "0.00000") + " " + Format(.v_13, "0.00000") + " " + Format(.v_14, "0.00000")
        Debug.Print Format(.v_21, "0.00000") + " " + Format(.v_22, "0.00000") + " " + Format(.v_23, "0.00000") + " " + Format(.v_24, "0.00000")
        Debug.Print Format(.v_31, "0.00000") + " " + Format(.v_32, "0.00000") + " " + Format(.v_33, "0.00000") + " " + Format(.v_34, "0.00000")
        Debug.Print Format(.v_41, "0.00000") + " " + Format(.v_42, "0.00000") + " " + Format(.v_43, "0.00000") + " " + Format(.v_44, "0.00000")
    End With
    
End Sub

' MLOOKAT: Calculates a transformation matrix for the view
Public Function MLookAt(dCamPosition As D3DVECTOR, dCamLookAt As D3DVECTOR) As D3DMATRIX

    ' Setup local variables ...
    
        Dim L_dVU As D3DVECTOR
        Dim L_dVR As D3DVECTOR
        Dim L_dVView As D3DVECTOR
        Dim L_dVDefaultUp As D3DVECTOR
        
        ' Set world up vector (y-Axis is up)
        L_dVDefaultUp.Y = -1
        
    ' Calculate camera transform ...
    
        ' Load result with identity matrix
        MIdentity MLookAt
        
        ' Calculate vector from position to look-at-point
        L_dVView = VNormalize(VSubtract(dCamLookAt, dCamPosition))
        
        ' Calculate right component of view vector
        L_dVR = VCrossProduct(L_dVDefaultUp, L_dVView)
        
        ' Calculate up component of view vector
        L_dVU = VCrossProduct(L_dVView, L_dVR)
        
        ' Normalize right and up
        L_dVR = VNormalize(L_dVR)
        L_dVU = VNormalize(L_dVU)
        
        ' Compose camera matrix
        With MLookAt
            .v_11 = L_dVR.X
            .v_21 = L_dVR.Y
            .v_31 = L_dVR.z
            .v_12 = L_dVU.X
            .v_22 = L_dVU.Y
            .v_32 = L_dVU.z
            .v_13 = L_dVView.X
            .v_23 = L_dVView.Y
            .v_33 = L_dVView.z
            .v_41 = -VDotProduct(L_dVR, dCamPosition)
            .v_42 = -VDotProduct(L_dVU, dCamPosition)
            .v_43 = -VDotProduct(L_dVView, dCamPosition)
        End With
    
End Function

' MPROJECT: Calculates a transformation matrix for the projection
Public Function MProject(nNear As Single, nFar As Single, nAngleFOV As Single) As D3DMATRIX

    ' Setup local variables ...
        Dim nC As Single
        Dim nS As Single
        Dim nQ As Single
        
    ' Calculate matrix ...
    
        nC = Cos(nAngleFOV * 0.5 * PIFactor)
        nS = Sin(nAngleFOV * 0.5 * PIFactor)
        nQ = nS / (1 - nNear / nFar)

        MIdentity MProject
        
        With MProject
            .v_11 = nC
            .v_22 = nC
            .v_33 = nQ
            .v_43 = -nQ * nNear
            .v_34 = nS
            .v_44 = 0
        End With

End Function

' VNORMALIZE: Returns normalized vector
Public Function VNormalize(dVector As D3DVECTOR) As D3DVECTOR
    VNormalize = dVector
    D3DRMVectorNormalize VNormalize
End Function

' VREFLECT: Calculates reflection of a vector upon a normal
Public Function VReflect(dRay As D3DVECTOR, dReflector As D3DVECTOR) As D3DVECTOR
    D3DRMVectorReflect VReflect, dRay, dReflector
End Function

' VCROSSPRODUCT: Calculates the cross product of two vectors
Public Function VCrossProduct(dV1 As D3DVECTOR, dV2 As D3DVECTOR) As D3DVECTOR
    D3DRMVectorCrossProduct VCrossProduct, dV1, dV2
End Function

' VDOTPRODUCT: Calculates the dot product of two vectors
Public Function VDotProduct(dV1 As D3DVECTOR, dV2 As D3DVECTOR) As Single
     VDotProduct = D3DRMVectorDotProduct(dV1, dV2)
End Function

' VSUBTRACT: Calculates the subtraction of V1 and V2
Public Function VSubtract(dV1 As D3DVECTOR, dV2 As D3DVECTOR) As D3DVECTOR
    VSubtract.X = dV1.X - dV2.X
    VSubtract.Y = dV1.Y - dV2.Y
    VSubtract.z = dV1.z - dV2.z
End Function

' VADD: Calculates the addition of V1 and V2
Public Function VAdd(dV1 As D3DVECTOR, dV2 As D3DVECTOR) As D3DVECTOR
    VAdd.X = dV1.X + dV2.X
    VAdd.Y = dV1.Y + dV2.Y
    VAdd.z = dV1.z + dV2.z
End Function

' VROTATE: Rotates a vector
Public Function VRotate(dV As D3DVECTOR, ByVal nValueX As Single, ByVal nValueY As Single, ByVal nValueZ As Single) As D3DVECTOR
    
    ' Setup local variables...
        Dim dVDir As D3DVECTOR  ' Holds axis for rotation
    
    ' Do rotations ...
        
        ' Rotate aroun X axis
        If nValueX <> 0 Then
            dVDir.X = 1
            dVDir.Y = 0
            dVDir.z = 0
            D3DRMVectorRotate VRotate, dV, dVDir, nValueX * PIFactor
        End If
    
        ' Rotate aroun Y axis
        If nValueY <> 0 Then
            dVDir.X = 0
            dVDir.Y = 1
            dVDir.z = 0
            D3DRMVectorRotate VRotate, dV, dVDir, nValueY * PIFactor
        End If
    
        ' Rotate aroun Z axis
        If nValueZ <> 0 Then
            dVDir.X = 0
            dVDir.Y = 0
            dVDir.z = 1
            D3DRMVectorRotate VRotate, dV, dVDir, nValueZ * PIFactor
        End If
    
End Function



