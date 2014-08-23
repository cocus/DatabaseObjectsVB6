VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSQLiteRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME                As String = "cSQLiteRecordset"


Private c_lSQLiteHandle                 As Long
Private c_lQuery                        As Long
Private c_bEOF                          As Boolean
Private c_bOpened                       As Boolean


Public Function LinkToQuery(ByVal lQuery As Long) As Boolean
    c_lQuery = lQuery
    c_bOpened = True
    
    #If SQLITE_DEBUG Then
        Debug.Print CLASS_NAME & "::LinkToQuery Handle: "; lQuery
    #End If
    Call MoveNext
End Function

Public Function LookUpColumnName(ByVal sName As String, _
                                 Optional ByRef lIndex As Long) As Boolean

    For lIndex = 0 To Me.FieldsCount()
        If StrComp(sName, sqlite3_column_name(c_lQuery, lIndex), vbTextCompare) = 0 Then
            LookUpColumnName = True
            Exit Function
        End If
    Next lIndex
End Function

'// Fields emulation
Public Property Get FieldsCount() As Long
    If Not c_bOpened Then
        RaiseError dboErrorGeneral, CLASS_NAME & "::FieldsCount Recordset is closed"
        Exit Property
    End If
    
    FieldsCount = sqlite3_column_count(c_lQuery)
End Property

'// Field(id) Emulation
Public Property Get FieldName(ByVal lIndex As Long) As String
    If (lIndex < 0) Or _
       (lIndex > Me.FieldsCount()) Then
        RaiseError dboErrorGeneral, CLASS_NAME & "::FieldName Column index is out of range"
        Exit Property
    End If
    
    FieldName = sqlite3_column_name(c_lQuery, lIndex)
End Property

Public Property Get FieldValue(ByVal lIndex As Long) As Variant
    If (lIndex < 0) Or _
       (lIndex > Me.FieldsCount()) Then
        RaiseError dboErrorGeneral, CLASS_NAME & "::FieldValue Column index is out of range"
        Exit Property
    End If

    FieldValue = sqlite3_column_text16(c_lQuery, lIndex)
End Property











Public Function MoveNext() As Boolean
    If Not c_bOpened Then
        RaiseError dboErrorGeneral, CLASS_NAME & "::MoveNext Recordset is closed"
        Exit Function
    End If
    
    MoveNext = (sqlite3_step(c_lQuery) = SQLITE_ROW)
    c_bEOF = Not MoveNext
End Function

Public Property Get EOF() As Boolean
    If Not c_bOpened Then
        RaiseError dboErrorGeneral, CLASS_NAME & "::EOF Recordset is closed"
        Exit Property
    End If
    
    EOF = c_bEOF
End Property

Public Function CloseRS() As Boolean
    If c_bOpened Then
        #If SQLITE_DEBUG Then
            Debug.Print CLASS_NAME & "::CloseRS Handle: "; c_lQuery
        #End If
        Call sqlite3_finalize(c_lQuery)
        c_bOpened = False
    End If
End Function




Private Sub Class_Terminate()
    Call CloseRS
End Sub
