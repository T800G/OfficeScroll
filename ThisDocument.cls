Option Explicit

Private Const S_OK = 0
Private hResult As Long

#If Win64 Then
Private Const DLLNAME = "MsoScroll64.dll"
Private Declare PtrSafe Function Connect Lib "MsoScroll64.dll" (ByVal r As Object) As Long
Private Declare PtrSafe Function Disconnect Lib "MsoScroll64.dll" () As Long
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private m_hDll As LongPtr
#Else
Private Const DLLNAME = "MsoScroll.dll"
Private Declare Function Connect Lib "MsoScroll.dll" (ByVal r As Object) As Long
Private Declare Function Disconnect Lib "MsoScroll.dll" () As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private m_hDll As Long
#End If

Public Sub AutoExec()
    Dim dllPath As String
    dllPath = Environ("APPDATA") & "\Microsoft\AddIns\" & DLLNAME
    If Dir(dllPath) <> "" Then
		'If Not CheckWDAddins() Then Exit Sub
        Debug.Assert (m_hDll = 0)
        If (m_hDll = 0) Then m_hDll = LoadLibrary(dllPath)
        Debug.Assert (m_hDll <> 0)
        If m_hDll <> 0 Then
            hResult = Connect(Application)
            Debug.Assert (hResult = S_OK)
			Debug.Print DLLNAME & "::Connect()"
        End If
    Else
        MsgBox DLLNAME & " file not found", vbCritical, "Office Scroll Add-in"
    End If
End Sub

Public Sub AutoExit()
    If (m_hDll <> 0) Then
        hResult = Disconnect()
        Debug.Assert (hResult = S_OK)
        FreeLibrary m_hDll
        m_hDll = 0
    End If
End Sub
