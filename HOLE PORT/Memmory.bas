Attribute VB_Name = "Memmory"
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpmstmemstat As MEMORYSTATUS)

Type MEMORYSTATUS
dwlength As Long       '32
dwmemoryload As Long   'percent of memory in use
dwtotalphys As Long    'bytes of phisical memory
dwavailphys As Long    'free phisical memory bytes
dwtotalpagefile As Long 'bytes of paging file
dwavailpagefile As Long 'free bytes of paging file
dwtotalvirtual As Long 'user bytes of address space
dwavailvirtual As Long 'free user bytes
End Type


Dim ms As MEMORYSTATUS


Public Function TotalMem()
TotalMem = ms.dwavailvirtual
End Function

Public Function AvailablePhysMem()
AvailablePhysMem = ms.dwavailphys
End Function

Public Function TotalPhysMem()
TotalPhysMem = ms.dwtotalphys
End Function

Public Function TotalVirtualMem()
TotalVirtualMem = ms.dwtotalvirtual
End Function

Public Function AvailablePageFile()
AvailablePageFile = ms.dwavailpagefile
End Function

Public Function TotalPageFile()
TotalPageFile = ms.dwtotalpagefile
End Function

Public Function MemoryLoad()
MemoryLoad = ms.dwmemoryload
End Function

Public Function Length1()
Length1 = ms.dwlength
End Function
