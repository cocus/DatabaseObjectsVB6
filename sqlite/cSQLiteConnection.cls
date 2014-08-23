VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSQLiteConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME                As String = "cSQLiteConnection"

Private c_lSQLiteHandle                 As Long


Public Function OpenFile(ByVal sFile As String) As Boolean
    If Not (sqlite3_open(sFile, c_lSQLiteHandle) = SQLITE_OK) Then
        RaiseError dboErrorGeneral, CLASS_NAME & ": Error loading database " & sFile & "!"
        Exit Function
    End If
    
    OpenFile = True
End Function


Public Function Execute(ByVal CommandText As String, _
                        Optional ByRef RecordsAffected As Long, _
                        Optional ByVal Options As Long = -1) As Object
    Dim lQuery                  As Long
    
    
    If (Options And adExecuteNoRecords) = adExecuteNoRecords Then
        '// No records, only execute SQL
        If Not (sqlite3_exec(c_lSQLiteHandle, CommandText) = SQLITE_OK) Then
            RaiseError dboErrorGeneral, CLASS_NAME & ": Error executing query: " & CommandText
            Exit Function
        End If
        
        RecordsAffected = sqlite3_changes(c_lSQLiteHandle)
    Else
        '// Return records
        If Not (sqlite3_prepare_v2(c_lSQLiteHandle, CommandText, 0, lQuery, 0) = SQLITE_OK) Then
            RaiseError dboErrorGeneral, CLASS_NAME & ": Error executing query: " & CommandText
            Exit Function
        End If
    
        RecordsAffected = sqlite3_changes(c_lSQLiteHandle)
        
        Set Execute = New cSQLiteRecordset
        Call Execute.LinkToQuery(lQuery)
    End If
End Function





Private Sub Class_Initialize()
    #If SQLITE_DEBUG Then
        Debug.Print CLASS_NAME & ": Initializing!"
    #End If
    
    If Not (sqlite3_initialize(App.Path) = SQLITE_OK) Then
        RaiseError dboErrorGeneral, CLASS_NAME & ": Error initializing library!"
    End If
End Sub

Private Sub Class_Terminate()
    #If SQLITE_DEBUG Then
        Debug.Print CLASS_NAME & ": Terminating! Handle: "; c_lSQLiteHandle
    #End If
    
    If c_lSQLiteHandle Then
        Call sqlite3_close(c_lSQLiteHandle)
    End If
    
    Call sqlite3_shutdown
End Sub

