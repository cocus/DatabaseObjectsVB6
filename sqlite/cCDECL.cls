VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCDECL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------------------------------------
'
' cCDECL - Class that enables the user to call cdecl dynamic link libraries.
'          Supports cdecl style variable argument lists and bas module
'          callbacks.
'
'031029 First cut....................................................... v1.00
'071129 Uses virtual memory rather than string space to fix a DEP issue. v1.01
'

Option Explicit
Option Base 0

'API declarations
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Long)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Byte)
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
Private Declare Sub PutMem8 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Currency)
Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

'Private constants
Private Const ERR_SRC       As String = "cCDECL"    'Error source name
Private Const ERR_NUM       As Long = vbObjectError 'cCDECL error number base
Private Const MAX_ARG       As Long = 16            'Maximum number of parameters, you can change this if required
Private Const PATCH_01      As Long = 15            'CDECL patch, CDECL function address
Private Const PATCH_02      As Long = 10            'Callback patch, bas mod function address patch
Private Const PATCH_03      As Long = 16            'Callback patch, stack adjustment patch

'Parameter block
Private Type tParamBlock
  ParamCount                As Long                 'Number of parameters to be forwarded to the cdecl function
  Params(0 To MAX_ARG - 1)  As Long                 'Array of parameters to be forwarded to the cdecl function
End Type

'Private member
Private m_LastError         As Long                 'Last error private member

'Private variables
Private bNewDLL             As Boolean              'Flag to indicate that the loaded DLL has changed
Private hMod                As Long                 'DLL module handle
Private nAddr               As Long                 'Cache the previous cdecl function's address
Private pCode               As Long                 'Pointer to the CDECL code
Private sLastFunc           As String               'Cache the previous cdecl function's name
Private pb                  As tParamBlock          'Parameter block instance

'Replace the stub proc (z_DO_NOT_CALL) with machine-code to handle the cdecl function
Private Sub Class_Initialize()
    Dim pMe As Long
  
    'Get the address of my vtable into pMe
    GetMem4 ObjPtr(Me), pMe

    'Allocate a page of executable memory
    pCode = VirtualAlloc(0, &H1000&, &H1000&, &H40&)
  
    'Copy the CDECL translation code to memory
    PutMem8 pCode + 0, -208642111809017.9757@
    PutMem8 pCode + 8, -605931634821031.5515@
    PutMem8 pCode + 16, 20765931315670.1386@
    PutMem8 pCode + 24, -857143604525899.4687@
    PutMem4 pCode + 32, &HC2C03102
    PutMem2 pCode + 36, &HC

    'Patch the first vtable entry (z_DO_NOT_CALL) to point to the CDECL code
    PutMem4 pMe + &H1C, pCode
  
    'Copy the callback thunk code to memory
    PutMem8 pCode + 40, 479615108421936.7656@
    PutMem8 pCode + 48, -140483859888551.3191@
    PutMem8 pCode + 56, 99649511.6971@
    PutMem8 pCode + 64, 21442817159.0144@
End Sub

Private Sub Class_Terminate()
    'Free virtual memory
    VirtualFree pCode, 0, &H8000&
End Sub

'This sub is replaced by machine code at pCode at class instance creation...
'IT MUST ONLY be called internally by CallFunc.
Public Function z_DO_NOT_CALL(ByVal nAddrParamBlock As Long) As Long
End Function

'Purpose:
' Call the named cdecl function with the passed parameters
'
'Arguments:
' sFunction - Name of the cdecl function to call
' ParmLongs - ParamArray of parameters to pass to the named cdecl function
'
'Return:
'  The return value of the named cdecl function
Public Function CallFunc(ByVal sFunction As String, ParamArray ParmLongs() As Variant) As Long
    Dim i As Long
    Dim j As Long
  
    'Check that the DLL is loaded
    If hMod = 0 Then
    
        'If in the IDE just stop, programmer may have On Error Resume Next protection and miss the problem.
        Debug.Assert False
        Call Err.Raise(ERR_NUM + 0, ERR_SRC, "DLL not loaded")
    End If

    'Check to see if we're calling the same cdecl function as the previous call to CallFunc
    If (StrComp(sLastFunc, sFunction) <> 0) Or bNewDLL Then
    
        'Get the address of the function
        nAddr = GetProcAddress(hMod, sFunction)
        If nAddr = 0 Then
      
            'If in the IDE just stop, programmer may have On Error Resume Next protection and miss the problem.
            Debug.Assert False
            Call Err.Raise(ERR_NUM + 1, ERR_SRC, "Failed to locate function: " & sFunction)
        End If

        'Patch the code buffer to call the relative address to the cdecl function
        PutMem4 pCode + PATCH_01, nAddr - pCode - (PATCH_01 + 4)
        bNewDLL = False
        sLastFunc = sFunction
    End If
  
    With pb
        j = UBound(ParmLongs)
        If j >= MAX_ARG Then
      
            'If in the IDE just stop, programmer may have On Error Resume Next protection and miss the problem.
            Debug.Assert False
            Call Err.Raise(ERR_NUM + 2, ERR_SRC, "Too many parameters")
        End If
    
        'Fill the parameter block
        For i = 0 To j
            .Params(i) = ParmLongs(i)
        Next i
    
        .ParamCount = i                                         '(j + 1)
    End With
  
    Call SetLastError(0)                                      'Clear the error code
    CallFunc = z_DO_NOT_CALL(VarPtr(pb))                      'Execute the code buffer passing the address of the parameter block
    m_LastError = GetLastError()                              'Get error code
End Function

'Load the DLL
Public Function DllLoad(ByVal sName As String) As Boolean
    hMod = LoadLibraryA(sName)

    If hMod <> 0 Then
        DllLoad = True
        'It's remotely possible that the programmer could change the dll and then call a function
        'in the new dll with exactly the same name as the previous CallFunc to the previous DLL. This would
        'defeat the caching scheme and result in the old function in the old dll being called. An unlikely
        'scenario, but stranger things have happened. Soooo, explicitly indicate that we're using a new dll
        bNewDLL = True
    End If
  
    'If in the IDE just stop on failure, programmer may not be checking the return value.
    Debug.Assert DllLoad
End Function

'It's not important to do this, but, if you've finished with a DLL there's no harm in releasing
'its memory. Don't bother at app end... it will be dealt with automatically when the process ends.
Public Function DllUnload() As Boolean
    If hMod <> 0 Then
        DllUnload = (FreeLibrary(hMod) <> 0)
        hMod = 0
    End If
  
    'If in the IDE, get the programmer's attention
    Debug.Assert DllUnload
End Function

'Return the cdecl function's error code
Public Property Get LastError() As Long
    LastError = m_LastError
End Property

'Purpose:
' Setup a wrapper so that a bas module function can act as a cdecl callback
'
'Arguments:
' nModFuncAddr - The address of the bas module function to act as a cdecl callback (use AddressOf)
' nParms       - The number of parameters that will be passed to the bas module function
'
'Return:
'  The address to pass to the cdecl function as the callback address
'
Public Function WrapCallback(ByVal nModFuncAddr As Long, ByVal nParms As Long) As Long
    Dim nStackAdjust As Long                                  'The number of bytes to adjust the stack
  
    WrapCallback = pCode + 40                                 'Address of the callback wrapper
    nStackAdjust = nParms * 4                                 'Four bytes per parameter

    'Patch the code buffer to call the vb bas module callback function
    PutMem4 WrapCallback + PATCH_02, nModFuncAddr - WrapCallback - (PATCH_02 + 4)
  
    'Patch the code buffer to apply the necessary stack adjustment
    PutMem4 WrapCallback + PATCH_03, nStackAdjust
End Function
