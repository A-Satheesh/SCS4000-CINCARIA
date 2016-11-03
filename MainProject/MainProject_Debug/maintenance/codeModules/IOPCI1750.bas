Attribute VB_Name = "IOPCI1750"
Public m_lDevHandle As Long
Global Const nDevNum = 0
Global Const nPortStart = 0

Public Function InitializePCI1750() As Boolean
    'Open the device
    'Call DRV_DeviceOpen(nDevNum, m_lDevHandle)
    
    Dim errCode As Long
    
    InitializePCI1750 = False
    
    errCode = DRV_DeviceOpen(nDevNum, m_lDevHandle)
    
    If (errCode = 0) Then
        InitializePCI1750 = True
    End If
End Function

Public Function Close_PCI1750() As Boolean
    'Close the device
    'Call DRV_DeviceClose(m_lDevHandle)
    
    Dim errCode As Long
    
    Close_PCI1750 = False
    
    errCode = DRV_DeviceClose(m_lDevHandle)
    
    If (errCode = 0) Then
        Close_PCI1750 = True
    End If
End Function

