Attribute VB_Name = "mDrivers"
Option Explicit
    
' Driver type for enumeration of D3D driver...
   
    ' Type to hold driver data
    Public Type tD3DDriver
        DESC    As String                         ' Driver description
        Name    As String                         ' Driver name
        GUID    As Byte                           ' Unique interface ID for accessing driver
        GUID1   As Byte                           ' ...
        GUID2   As Byte                           ' ...
        GUID3   As Byte                           ' ...
        GUID4   As Byte                           ' ...
        GUID5   As Byte                           ' ...
        GUID6   As Byte                           ' ...
        GUID7   As Byte                           ' ...
        GUID8   As Byte                           ' ...
        GUID9   As Byte                           ' ...
        GUID10  As Byte                           ' ...
        GUID11  As Byte                           ' ...
        GUID12  As Byte                           ' ...
        GUID13  As Byte                           ' ...
        GUID14  As Byte                           ' ...
        GUID15  As Byte                           ' ...
        DEVDESC As D3DDEVICEDESC                  ' Device description for use by D3D
    End Type
    
' DriverSelected enum: Tells which DD/D3D driver to use
    Public Enum eDXDriverType
        EDXDTSoft = 0
        EDXDTHard = 1
        EDXDTPlus = 2
    End Enum
    
' Driver type for enumeration of DD driver...
    Public Type tDDDriver
        DESC    As String                         ' Driver description
        Name    As String                         ' Driver name
        GUID    As Byte                           ' Unique interface ID for accessing driver
        GUID1   As Byte                           ' ...
        GUID2   As Byte                           ' ...
        GUID3   As Byte                           ' ...
        GUID4   As Byte                           ' ...
        GUID5   As Byte                           ' ...
        GUID6   As Byte                           ' ...
        GUID7   As Byte                           ' ...
        GUID8   As Byte                           ' ...
        GUID9   As Byte                           ' ...
        GUID10  As Byte                           ' ...
        GUID11  As Byte                           ' ...
        GUID12  As Byte                           ' ...
        GUID13  As Byte                           ' ...
        GUID14  As Byte                           ' ...
        GUID15  As Byte                           ' ...
        D3DDriver As tD3DDriver                   ' Subordinate software driver
        DriverType As eDXDriverType               ' Driver Type
        Found As Boolean                          ' Driver status
    End Type
    
' ENUMD3DEVICECALLBACK: Enumerates Device drivers for Direct3D
Public Function EnumD3DDeviceCallback(nGUID As Long, nDDDesc As Long, nDDName As Long, dHALD3DDevDesc As D3DDEVICEDESC, dHELD3DDevDesc As D3DDEVICEDESC, nOptions As Long) As Long

    ' Enable error handling ...
        On Error Resume Next
    
    ' Setup local variables ...
    
        Dim L_nTemp(256) As Byte                      ' Temporary array for name and guid translation
        Dim L_nChar As Byte                           ' Temporary charactar for name translation
        Dim L_nIndex As Long                          ' Variable to run through temp array
        Dim L_bHDW As Boolean                         ' Flag for hardware support
        Dim L_dD3DDEVDESC As D3DDEVICEDESC            ' Temporary device description
        Dim L_dD3DDriver As tD3DDriver                ' Temporary holds driver description
        
    ' Get driver capabilities...
      
        ' Decide if hardware supports rgb color model and enable HAL or HEL support properly
        L_bHDW = (dHALD3DDevDesc.dcmColorModel = D3DCOLOR_RGB)
        
        ' Get device description from HAL if hardware support or from HEL if software renderer
         If L_bHDW Then
            L_dD3DDEVDESC = dHALD3DDevDesc
         Else
            L_dD3DDEVDESC = dHELD3DDevDesc
         End If
         
    ' Decide if driver fits application needs ...
        
        ' Do not use hardware drivers with software rendering ...
        If G_dDXSelectedDriver.DriverType = EDXDTSoft And L_bHDW Then
            EnumD3DDeviceCallback = DDENUMRET_OK
            Exit Function
        End If
        
        ' Do not use software drivers with hardware rendering  ...
        If G_dDXSelectedDriver.DriverType = EDXDTHard And Not L_bHDW Then
            EnumD3DDeviceCallback = DDENUMRET_OK
            Exit Function
        End If
        
        ' Do not use software drivers with addon boards ...
        If G_dDXSelectedDriver.DriverType = EDXDTPlus And Not L_bHDW Then
            EnumD3DDeviceCallback = DDENUMRET_OK
            Exit Function
        End If
        
        ' Do not use MONO RAMP devices at all
        If Not (L_dD3DDEVDESC.dcmColorModel = D3DCOLOR_RGB) Then
            EnumD3DDeviceCallback = DDENUMRET_OK
            Exit Function
        End If
    
        ' Do not use devices having render bit depth below 8 bit
        If ((L_dD3DDEVDESC.dwDeviceRenderBitDepth And DDBD_8) = 0) And ((L_dD3DDEVDESC.dwDeviceRenderBitDepth And DDBD_16) = 0) And ((L_dD3DDEVDESC.dwDeviceRenderBitDepth And DDBD_24) = 0) And ((L_dD3DDEVDESC.dwDeviceRenderBitDepth And DDBD_32) = 0) Then
            EnumD3DDeviceCallback = DDENUMRET_OK
            Exit Function
        End If
            
    ' DRIVER ACCEPTED ...
    
        ' Get driver info ...
        With L_dD3DDriver
            
            ' Set driver data description
            .DEVDESC = L_dD3DDEVDESC
            
            ' Copy GUID data into temporary array
            CopyMemory VarPtr(L_nTemp(0)), VarPtr(nGUID), 16
            
            ' Set GUID data into driver structure
            .GUID = L_nTemp(0)
            .GUID1 = L_nTemp(1)
            .GUID2 = L_nTemp(2)
            .GUID3 = L_nTemp(3)
            .GUID4 = L_nTemp(4)
            .GUID5 = L_nTemp(5)
            .GUID6 = L_nTemp(6)
            .GUID7 = L_nTemp(7)
            .GUID8 = L_nTemp(8)
            .GUID9 = L_nTemp(9)
            .GUID10 = L_nTemp(10)
            .GUID11 = L_nTemp(11)
            .GUID12 = L_nTemp(12)
            .GUID13 = L_nTemp(13)
            .GUID14 = L_nTemp(14)
            .GUID15 = L_nTemp(15)
            
            ' Copy driver name into temporary array
            CopyMemory VarPtr(L_nTemp(0)), VarPtr(nDDName), 255
              
            ' Parse name of driver
            For L_nIndex = 0 To 255
                L_nChar = L_nTemp(L_nIndex)
                If L_nChar < 32 Then Exit For
                .Name = .Name + Chr(L_nChar)
            Next
            
            ' Copy driver Description into temporary array
            CopyMemory VarPtr(L_nTemp(0)), VarPtr(nDDDesc), 255
              
            ' Parse description of driver
            For L_nIndex = 0 To 255
                L_nChar = L_nTemp(L_nIndex)
                If L_nChar < 32 Then Exit For
                .DESC = .DESC + Chr(L_nChar)
            Next
            
        End With
        
        ' Set driver info ...
        Select Case G_dDXSelectedDriver.DriverType
            Case EDXDTSoft
                G_dDXDriverSoft.D3DDriver = L_dD3DDriver
                G_dDXDriverSoft.Found = True
            Case EDXDTHard
                G_dDXDriverHard.D3DDriver = L_dD3DDriver
                G_dDXDriverHard.Found = True
            Case EDXDTPlus
                G_dDXDriverPlus.D3DDriver = L_dD3DDriver
                G_dDXDriverPlus.Found = True
        End Select
        
        ' Return success
        EnumD3DDeviceCallback = DDENUMRET_OK
    
    ' Error handling ...
        Exit Function
    
End Function

' ENUMDDDEVICECALLBACK: Enumerates DirectDraw drivers and selects the best one
Public Function EnumDDDeviceCallback(nGUID As Long, nDDDesc As Long, nDDName As Long, nOptions As Long) As Long

    ' Enable error handling ...
    
        On Error Resume Next
        
    ' Setup local variables ...
    
        Dim L_nTemp(256) As Byte                      ' Temporary array for name and guid translation
        Dim L_nChar As Byte                           ' Temporary charactar for name translation
        Dim L_nIndex As Long                          ' Variable to run through temp array
        Dim L_dDDDriver As tDDDriver                  ' Temporary holds driver data
        
    ' Process current driver ...
            
        ' Get driver info ...
                    
        With L_dDDDriver
        
            ' Copy GUID data into temporary array
            CopyMemory VarPtr(L_nTemp(0)), VarPtr(nGUID), 16
            
            ' Set GUID data into driver structure
            .GUID = L_nTemp(0)
            .GUID1 = L_nTemp(1)
            .GUID2 = L_nTemp(2)
            .GUID3 = L_nTemp(3)
            .GUID4 = L_nTemp(4)
            .GUID5 = L_nTemp(5)
            .GUID6 = L_nTemp(6)
            .GUID7 = L_nTemp(7)
            .GUID8 = L_nTemp(8)
            .GUID9 = L_nTemp(9)
            .GUID10 = L_nTemp(10)
            .GUID11 = L_nTemp(11)
            .GUID12 = L_nTemp(12)
            .GUID13 = L_nTemp(13)
            .GUID14 = L_nTemp(14)
            .GUID15 = L_nTemp(15)
            
            ' Copy driver name into temporary array
            CopyMemory VarPtr(L_nTemp(0)), VarPtr(nDDName), 255
              
            ' Parse name of driver
            For L_nIndex = 0 To 255
                L_nChar = L_nTemp(L_nIndex)
                If L_nChar < 32 Then Exit For
                .Name = .Name + Chr(L_nChar)
            Next
            
            ' Copy driver Description into temporary array
            CopyMemory VarPtr(L_nTemp(0)), VarPtr(nDDDesc), 255
              
            ' Parse description of driver
            For L_nIndex = 0 To 255
                L_nChar = L_nTemp(L_nIndex)
                If L_nChar < 32 Then Exit For
                .DESC = .DESC + Chr(L_nChar)
            Next
                    
        End With
        
        ' Set driver info ...
        If G_bPrimaryDisplayAlreadyDetected Then
            G_dDXDriverPlus = L_dDDDriver
            G_dDXDriverPlus.Found = True
        Else
            G_dDXDriverHard = L_dDDDriver
            G_dDXDriverSoft = L_dDDDriver
            G_bPrimaryDisplayAlreadyDetected = True
        End If
            
        ' Return success
        EnumDDDeviceCallback = DDENUMRET_OK

    
    ' Error handling ...
        Exit Function
        
End Function

