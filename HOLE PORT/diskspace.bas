Attribute VB_Name = "DiskSpace"
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
    
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

    Dim lngSectorsPerCluster As Long
    Dim lngBytesPerSector As Long
    Dim lngNumberOfFreeClusters As Long
    Dim lngTotalNumberOfClusters As Long
    Dim strBuffer As String * 1024
    
Public Function GetDiskFreeSpace1()
        
        RetVal = GetDiskFreeSpace("c:\", lngSectorsPerCluster, lngBytesPerSector, lngNumberOfFreeClusters, lngTotalNumberOfClusters)

        GetDiskFreeSpace1 = RetVal
End Function
    
    
Public Function GetWindowsDirectory1()
    retval2 = GetWindowsDirectory(strBuffer, 1024)
    GetWindowsDirectory1 = retval2
End Function
