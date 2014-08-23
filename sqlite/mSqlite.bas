Attribute VB_Name = "mSqlite"
'---------------------------------------------------------------------------------------
' Module      : mSqlite
' DateTime    : 10/11/2009 22:05
' Author      : Cobein
' Mail        : cobein27@hotmail.com
' WebPage     : http://www.vbsqlite.com.ar
' Purpose     :
' Usage       : At your own risk
' Requirements: None
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,-
'               online service, or distribute as source
'               on any media without express permission.
'
' Version     : 10122009
'
' History     : 10/11/2009 First Cut....................................................
'---------------------------------------------------------------------------------------
Option Explicit

Public Type DATA_BLOB
    cbData              As Long
    pbData              As Long
End Type

Public Const CP_UTF8                As Long = 65001

Public Const MEM_RELEASE            As Long = &H8000
Public Const MEM_COMMIT             As Long = &H1000
Public Const PAGE_READWRITE         As Long = &H4

Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Public Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function PathCanonicalize Lib "shlwapi.dll" Alias "PathCanonicalizeA" (ByVal pszBuf As String, ByVal pszPath As String) As Long

'// Compile-Time Library Version Numbers
Public Const SQLITE_VERSION         As String = "3.6.16"
Public Const SQLITE_VERSION_NUMBER  As Long = 3006016

'// Default library name
Public Const SQLITE3_LIB            As String = "sqlite3.dll"

'// Result Codes
Public Const SQLITE_OK              As Long = 0   '// Successful result
'// beginning-of-error-codes
Public Const SQLITE_ERROR           As Long = 1   '// SQL error or missing database
Public Const SQLITE_INTERNAL        As Long = 2   '// Internal logic error in SQLite
Public Const SQLITE_PERM            As Long = 3   '// Access permission denied
Public Const SQLITE_ABORT           As Long = 4   '// Callback routine requested an abort
Public Const SQLITE_BUSY            As Long = 5   '// The database file is locked
Public Const SQLITE_LOCKED          As Long = 6   '// A table in the database is locked
Public Const SQLITE_NOMEM           As Long = 7   '// A malloc() failed
Public Const SQLITE_READONLY        As Long = 8   '// Attempt to write a readonly database
Public Const SQLITE_INTERRUPT       As Long = 9   '// Operation terminated by sqlite3_interrupt()
Public Const SQLITE_IOERR           As Long = 10  '// Some kind of disk I/O error occurred
Public Const SQLITE_CORRUPT         As Long = 11  '// The database disk image is malformed
Public Const SQLITE_NOTFOUND        As Long = 12  '// NOT USED. Table or record not found
Public Const SQLITE_FULL            As Long = 13  '// Insertion failed because database is full
Public Const SQLITE_CANTOPEN        As Long = 14  '// Unable to open the database file
Public Const SQLITE_PROTOCOL        As Long = 15  '// NOT USED. Database lock protocol error
Public Const SQLITE_EMPTY           As Long = 16  '// Database is empty
Public Const SQLITE_SCHEMA          As Long = 17  '// The database schema changed
Public Const SQLITE_TOOBIG          As Long = 18  '// String or BLOB exceeds size limit
Public Const SQLITE_CONSTRAINT      As Long = 19  '// Abort due to constraint violation
Public Const SQLITE_MISMATCH        As Long = 20  '// Data type mismatch
Public Const SQLITE_MISUSE          As Long = 21  '// Library used incorrectly
Public Const SQLITE_NOLFS           As Long = 22  '// Uses OS features not supported on host
Public Const SQLITE_AUTH            As Long = 23  '// Authorization denied
Public Const SQLITE_FORMAT          As Long = 24  '// Auxiliary database format error
Public Const SQLITE_RANGE           As Long = 25  '// 2nd parameter to sqlite3_bind out of range
Public Const SQLITE_NOTADB          As Long = 26  '// File opened that is not a database file
Public Const SQLITE_ROW             As Long = 100 '// sqlite3_step() has another row ready
Public Const SQLITE_DONE            As Long = 101 '// sqlite3_step() has finished executing

'//Fundamental Datatypes
Public Const SQLITE_INTEGER         As Long = 1
Public Const SQLITE_FLOAT           As Long = 2
Public Const SQLITE3_TEXT           As Long = 3
Public Const SQLITE_BLOB            As Long = 4
Public Const SQLITE_NULL            As Long = 5

'// Run-Time Limit Categories
Public Const SQLITE_LIMIT_LENGTH                As Long = 0 '// The maximum size of any string or BLOB or table row.
Public Const SQLITE_LIMIT_SQL_LENGTH            As Long = 1 '// The maximum length of an SQL statement.
Public Const SQLITE_LIMIT_COLUMN                As Long = 2 '// The maximum number of columns in a table definition or in the result set of a SELECT or the maximum number of columns in an index or in an ORDER BY or GROUP BY clause.
Public Const SQLITE_LIMIT_EXPR_DEPTH            As Long = 3 '// The maximum depth of the parse tree on any expression.
Public Const SQLITE_LIMIT_COMPOUND_SELECT       As Long = 4 '// The maximum number of terms in a compound SELECT statement.
Public Const SQLITE_LIMIT_VDBE_OP               As Long = 5 '// The maximum number of instructions in a virtual machine program used to implement an SQL statement.
Public Const SQLITE_LIMIT_FUNCTION_ARG          As Long = 6 '// The maximum number of arguments on a function.
Public Const SQLITE_LIMIT_ATTACHED              As Long = 7 '// The maximum number of attached databases.
Public Const SQLITE_LIMIT_LIKE_PATTERN_LENGTH   As Long = 8 '// The maximum length of the pattern argument to the LIKE or GLOB operators.
Public Const SQLITE_LIMIT_VARIABLE_NUMBER       As Long = 9 '// The maximum number of variables in an SQL statement that can be bound.

'// Flags For File Open Operations
Public Const SQLITE_OPEN_READONLY               As Long = &H1     '// Ok for sqlite3_open_v2() */
Public Const SQLITE_OPEN_READWRITE              As Long = &H2     '// Ok for sqlite3_open_v2() */
Public Const SQLITE_OPEN_CREATE                 As Long = &H4     '// Ok for sqlite3_open_v2() */
Public Const SQLITE_OPEN_DELETEONCLOSE          As Long = &H8     '// VFS only */
Public Const SQLITE_OPEN_EXCLUSIVE              As Long = &H10    '// VFS only */
Public Const SQLITE_OPEN_MAIN_DB                As Long = &H100   '// VFS only */
Public Const SQLITE_OPEN_TEMP_DB                As Long = &H200   '// VFS only */
Public Const SQLITE_OPEN_TRANSIENT_DB           As Long = &H400   '// VFS only */
Public Const SQLITE_OPEN_MAIN_JOURNAL           As Long = &H800   '// VFS only */
Public Const SQLITE_OPEN_TEMP_JOURNAL           As Long = &H1000  '// VFS only */
Public Const SQLITE_OPEN_SUBJOURNAL             As Long = &H2000  '// VFS only */
Public Const SQLITE_OPEN_MASTER_JOURNAL         As Long = &H4000  '// VFS only */
Public Const SQLITE_OPEN_NOMUTEX                As Long = &H8000  '// Ok for sqlite3_open_v2() */
Public Const SQLITE_OPEN_FULLMUTEX              As Long = &H10000 '// Ok for sqlite3_open_v2() */


'// Authorizer Action Codes
'/**************************************************************** 3rd ************ 4th ***********/
Public Const SQLITE_CREATE_INDEX                As Long = 1   '// Index Name      Table Name      */
Public Const SQLITE_CREATE_TABLE                As Long = 2   '// Table Name      NULL            */
Public Const SQLITE_CREATE_TEMP_INDEX           As Long = 3   '// Index Name      Table Name      */
Public Const SQLITE_CREATE_TEMP_TABLE           As Long = 4   '// Table Name      NULL            */
Public Const SQLITE_CREATE_TEMP_TRIGGER         As Long = 5   '// Trigger Name    Table Name      */
Public Const SQLITE_CREATE_TEMP_VIEW            As Long = 6   '// View Name       NULL            */
Public Const SQLITE_CREATE_TRIGGER              As Long = 7   '// Trigger Name    Table Name      */
Public Const SQLITE_CREATE_VIEW                 As Long = 8   '// View Name       NULL            */
Public Const SQLITE_DELETE                      As Long = 9   '// Table Name      NULL            */
Public Const SQLITE_DROP_INDEX                  As Long = 10  '// Index Name      Table Name      */
Public Const SQLITE_DROP_TABLE                  As Long = 11  '// Table Name      NULL            */
Public Const SQLITE_DROP_TEMP_INDEX             As Long = 12  '// Index Name      Table Name      */
Public Const SQLITE_DROP_TEMP_TABLE             As Long = 13  '// Table Name      NULL            */
Public Const SQLITE_DROP_TEMP_TRIGGER           As Long = 14  '// Trigger Name    Table Name      */
Public Const SQLITE_DROP_TEMP_VIEW              As Long = 15  '// View Name       NULL            */
Public Const SQLITE_DROP_TRIGGER                As Long = 16  '// Trigger Name    Table Name      */
Public Const SQLITE_DROP_VIEW                   As Long = 17  '// View Name       NULL            */
Public Const SQLITE_INSERT                      As Long = 18  '// Table Name      NULL            */
Public Const SQLITE_PRAGMA                      As Long = 19  '// Pragma Name     1st arg or NULL */
Public Const SQLITE_READ                        As Long = 20  '// Table Name      Column Name     */
Public Const SQLITE_SELECT                      As Long = 21  '// NULL            NULL            */
Public Const SQLITE_TRANSACTION                 As Long = 22  '// Operation       NULL            */
Public Const SQLITE_UPDATE                      As Long = 23  '// Table Name      Column Name     */
Public Const SQLITE_ATTACH                      As Long = 24  '// Filename        NULL            */
Public Const SQLITE_DETACH                      As Long = 25  '// Database Name   NULL            */
Public Const SQLITE_ALTER_TABLE                 As Long = 26  '// Database Name   Table Name      */
Public Const SQLITE_REINDEX                     As Long = 27  '// Index Name      NULL            */
Public Const SQLITE_ANALYZE                     As Long = 28  '// Table Name      NULL            */
Public Const SQLITE_CREATE_VTABLE               As Long = 29  '// Table Name      Module Name     */
Public Const SQLITE_DROP_VTABLE                 As Long = 30  '// Table Name      Module Name     */
Public Const SQLITE_FUNCTION                    As Long = 31  '// NULL            Function Name   */
Public Const SQLITE_SAVEPOINT                   As Long = 32  '// Operation       Savepoint Name  */
Public Const SQLITE_COPY                        As Long = 0   '// No longer used */

'// Authorizer Return Codes
Public Const SQLITE_DENY                        As Long = 1   '// Abort the SQL statement with an error */
Public Const SQLITE_IGNORE                      As Long = 2   '// Don't allow access, but don't generate an error */

'// Mutex Types
Public Const SQLITE_MUTEX_FAST                  As Long = 0
Public Const SQLITE_MUTEX_RECURSIVE             As Long = 1
Public Const SQLITE_MUTEX_STATIC_MASTER         As Long = 2
Public Const SQLITE_MUTEX_STATIC_MEM            As Long = 3    '//  sqlite3_malloc() */
Public Const SQLITE_MUTEX_STATIC_MEM2           As Long = 4    '//  NOT USED */
Public Const SQLITE_MUTEX_STATIC_OPEN           As Long = 4    '//  sqlite3BtreeOpen() */
Public Const SQLITE_MUTEX_STATIC_PRNG           As Long = 5    '//  sqlite3_random() */
Public Const SQLITE_MUTEX_STATIC_LRU            As Long = 6    '//  lru page list */
Public Const SQLITE_MUTEX_STATIC_LRU2           As Long = 7    '//  lru page list */

'// File Locking Levels
Public Const SQLITE_LOCK_NONE           As Long = 0
Public Const SQLITE_LOCK_SHARED         As Long = 1
Public Const SQLITE_LOCK_RESERVED       As Long = 2
Public Const SQLITE_LOCK_PENDING        As Long = 3
Public Const SQLITE_LOCK_EXCLUSIVE      As Long = 4

'// Device Characteristics
Public Const SQLITE_IOCAP_ATOMIC        As Long = &H1
Public Const SQLITE_IOCAP_ATOMIC512     As Long = &H2
Public Const SQLITE_IOCAP_ATOMIC1K      As Long = &H4
Public Const SQLITE_IOCAP_ATOMIC2K      As Long = &H8
Public Const SQLITE_IOCAP_ATOMIC4K      As Long = &H10
Public Const SQLITE_IOCAP_ATOMIC8K      As Long = &H20
Public Const SQLITE_IOCAP_ATOMIC16K     As Long = &H40
Public Const SQLITE_IOCAP_ATOMIC32K     As Long = &H80
Public Const SQLITE_IOCAP_ATOMIC64K     As Long = &H100
Public Const SQLITE_IOCAP_SAFE_APPEND   As Long = &H200
Public Const SQLITE_IOCAP_SEQUENTIAL    As Long = &H400

Private m_cCDECL    As cCDECL
Private m_bLoaded   As Boolean

'---------------------------------------------------------------------------------------
' Procedure : sqlite3_initialize
' Purpose   : Wraper initialization
' Params    : [in] sPath: Optional directory path to sqlite library.
'             [out] sqlite3_initialize: sqlite status code.
'---------------------------------------------------------------------------------------
Public Function sqlite3_initialize(Optional ByVal sPath As String) As Long
     
    If m_bLoaded Then
        Call sqlite3_shutdown
    End If
    
    Set m_cCDECL = New cCDECL
       
    If sPath = vbNullString Then
        sPath = App.Path
    End If
    
    sPath = sPath & IIf(Right(sPath, 1) = "\", vbNullString, "\")
    sPath = Canonicalize(sPath)
    
    If m_cCDECL.DllLoad(sPath & SQLITE3_LIB) Then
        With m_cCDECL
            If .CallFunc("sqlite3_initialize") = SQLITE_OK Then
                sqlite3_initialize = SQLITE_OK
                m_bLoaded = True
            End If
        End With
    Else
        sqlite3_initialize = SQLITE_ERROR
        Call sqlite3_shutdown
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : sqlite3_shutdown
' Purpose   : Wraper termination
' Params    : [out] sqlite3_shutdown: sqlite status code.
'---------------------------------------------------------------------------------------
Public Function sqlite3_shutdown() As Long

    If Not m_cCDECL Is Nothing Then
        With m_cCDECL
            If .CallFunc("sqlite3_shutdown") = SQLITE_OK Then
                .DllUnload
                Set m_cCDECL = Nothing
                m_bLoaded = False
                sqlite3_shutdown = SQLITE_OK
                Exit Function
            End If
        End With
    End If
    
    sqlite3_shutdown = SQLITE_ERROR
  
End Function

Public Function sqlite3_open(ByVal filename As String, ByRef ppDb As Long) As Long
    'int sqlite3_open(
    '  const char *filename,   /* Database filename (UTF-8) */
    '  sqlite3 **ppDb          /* OUT: SQLite db handle */
    ');
    Dim bvStr() As Byte
    
    filename = UnicodeToUTF8(filename)
    
    bvStr = StrConv(filename & vbNullChar, vbFromUnicode)
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_open = .CallFunc("sqlite3_open", VarPtr(bvStr(0)), VarPtr(ppDb))
        End With
    Else
        sqlite3_open = SQLITE_ERROR
    End If

End Function

Public Function sqlite3_open16(ByVal filename As String, ByRef ppDb As Long) As Long
    'int sqlite3_open16(
    '  const void *filename,   /* Database filename (UTF-16) */
    '  sqlite3 **ppDb          /* OUT: SQLite db handle */
    ');
    Dim bvStr() As Byte
    
    filename = UnicodeToUTF16(filename)
    
    bvStr = StrConv(filename & vbNullChar, vbFromUnicode)
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_open16 = .CallFunc("sqlite3_open16", VarPtr(bvStr(0)), VarPtr(ppDb))
        End With
    Else
        sqlite3_open16 = SQLITE_ERROR
    End If

End Function

Public Function sqlite3_open_v2(ByVal filename As String, ByRef ppDb As Long, ByVal flags As Long, ByVal zVfs As String) As Long
    'int sqlite3_open_v2(
    '  const char *filename,   /* Database filename (UTF-8) */
    '  sqlite3 **ppDb,         /* OUT: SQLite db handle */
    '  int flags,              /* Flags */
    '  const char *zVfs        /* Name of VFS module to use */
    ');

    Dim bvStr()     As Byte
    Dim bvStr1()    As Byte
    
    filename = UnicodeToUTF8(filename)
    bvStr = StrConv(filename & vbNullChar, vbFromUnicode)
    
    zVfs = UnicodeToUTF8(zVfs)
    bvStr1 = StrConv(zVfs & vbNullChar, vbFromUnicode)
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_open_v2 = .CallFunc("sqlite3_open_v2", _
               VarPtr(bvStr(0)), VarPtr(ppDb), flags, VarPtr(bvStr1(0)))
        End With
    Else
        sqlite3_open_v2 = SQLITE_ERROR
    End If

End Function

Public Function sqlite3_close(ByVal sqlite3 As Long) As Long
    'int sqlite3_close(sqlite3 *);
        
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_close = .CallFunc("sqlite3_close", sqlite3)
        End With
    Else
        sqlite3_close = SQLITE_ERROR
    End If
End Function


Public Function sqlite3_changes(ByVal sqlite3 As Long) As Long
    'int sqlite3_changes(sqlite3 *);
        
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_changes = .CallFunc("sqlite3_changes", sqlite3)
        End With
    Else
        sqlite3_changes = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_prepare( _
       ByVal sqlite3 As Long, _
       ByVal zSql As String, _
       ByVal nByte As Long, _
       ByRef ppStmt As Long, _
       ByRef pzTail As Long) As Long
       
    'int sqlite3_prepare(
    '  sqlite3 *db,            /* Database handle */
    '  const char *zSql,       /* SQL statement, UTF-8 encoded */
    '  int nByte,              /* Maximum length of zSql in bytes. */
    '  sqlite3_stmt **ppStmt,  /* OUT: Statement handle */
    '  const char **pzTail     /* OUT: Pointer to unused portion of zSql */
    ');
    
    Dim bvStr() As Byte
    
    zSql = UnicodeToUTF8(zSql)
    bvStr = StrConv(zSql & vbNullChar, vbFromUnicode)
    
    If nByte = 0 Then
        nByte = Len(zSql)
    End If
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_prepare = .CallFunc( _
               "sqlite3_prepare", _
               sqlite3, _
               VarPtr(bvStr(0)), _
               nByte, _
               VarPtr(ppStmt), _
               VarPtr(pzTail))
        End With
    Else
        sqlite3_prepare = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_prepare16( _
       ByVal sqlite3 As Long, _
       ByVal zSql As String, _
       ByVal nByte As Long, _
       ByRef ppStmt As Long, _
       ByRef pzTail As Long) As Long
       
    'int sqlite3_prepare16(
    '  sqlite3 *db,            /* Database handle */
    '  const void *zSql,       /* SQL statement, UTF-16 encoded */
    '  int nByte,              /* Maximum length of zSql in bytes. */
    '  sqlite3_stmt **ppStmt,  /* OUT: Statement handle */
    '  const void **pzTail     /* OUT: Pointer to unused portion of zSql */
    ');

    Dim bvStr() As Byte
    
    zSql = UnicodeToUTF16(zSql)
    bvStr = StrConv(zSql & vbNullChar, vbFromUnicode)
    
    If nByte = 0 Then
        nByte = Len(zSql)
    End If
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_prepare16 = .CallFunc( _
               "sqlite3_prepare16", _
               sqlite3, _
               VarPtr(bvStr(0)), _
               nByte, _
               VarPtr(ppStmt), _
               VarPtr(pzTail))
        End With
    Else
        sqlite3_prepare16 = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_prepare_v2( _
       ByVal sqlite3 As Long, _
       ByVal zSql As String, _
       ByVal nByte As Long, _
       ByRef ppStmt As Long, _
       ByRef pzTail As Long) As Long
       
    'int sqlite3_prepare_v2(
    '  sqlite3 *db,            /* Database handle */
    '  const char *zSql,       /* SQL statement, UTF-8 encoded */
    '  int nByte,              /* Maximum length of zSql in bytes. */
    '  sqlite3_stmt **ppStmt,  /* OUT: Statement handle */
    '  const char **pzTail     /* OUT: Pointer to unused portion of zSql */
    ');
    
    Dim bvStr() As Byte
    
    zSql = UnicodeToUTF8(zSql)
    bvStr = StrConv(zSql & vbNullChar, vbFromUnicode)
    
    If nByte = 0 Then
        nByte = Len(zSql)
    End If
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_prepare_v2 = .CallFunc( _
               "sqlite3_prepare_v2", _
               sqlite3, _
               VarPtr(bvStr(0)), _
               nByte, _
               VarPtr(ppStmt), _
               VarPtr(pzTail))
        End With
        'SQLite_SetLastStatment ppStmt
    Else
        sqlite3_prepare_v2 = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_prepare16_v2( _
       ByVal sqlite3 As Long, _
       ByVal zSql As String, _
       ByVal nByte As Long, _
       ByRef ppStmt As Long, _
       ByRef pzTail As Long) As Long
       
    'int sqlite3_prepare16_v2(
    '  sqlite3 *db,            /* Database handle */
    '  const void *zSql,       /* SQL statement, UTF-16 encoded */
    '  int nByte,              /* Maximum length of zSql in bytes. */
    '  sqlite3_stmt **ppStmt,  /* OUT: Statement handle */
    '  const void **pzTail     /* OUT: Pointer to unused portion of zSql */
    ');

    Dim bvStr() As Byte
    
    zSql = UnicodeToUTF16(zSql)
    bvStr = StrConv(zSql & vbNullChar, vbFromUnicode)
    
    If nByte = 0 Then
        nByte = Len(zSql)
    End If
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_prepare16_v2 = .CallFunc( _
               "sqlite3_prepare16_v2", _
               sqlite3, _
               VarPtr(bvStr(0)), _
               nByte, _
               VarPtr(ppStmt), _
               VarPtr(pzTail))
        End With
    Else
        sqlite3_prepare16_v2 = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_exec(ByVal sqlite3 As Long, ByVal zSql As String) As Long
    'int sqlite3_exec(
    '  sqlite3*,                                  /* An open database */
    '  const char *sql,                           /* SQL to be evaluated */
    '  int (*callback)(void*,int,char**,char**),  /* Callback function */
    '  void *,                                    /* 1st argument to callback */
    '  char **errmsg                              /* Error msg written here */
    ');
    
    Dim bvStr() As Byte
    
    zSql = UnicodeToUTF8(zSql)
    bvStr = StrConv(zSql & vbNullChar, vbFromUnicode)

    If m_bLoaded Then
        With m_cCDECL
            sqlite3_exec = .CallFunc("sqlite3_exec", sqlite3, VarPtr(bvStr(0)), 0, 0, 0)
        End With
    Else
        sqlite3_exec = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_step(ByVal sqlite3_stmt As Long) As Long
    'int sqlite3_step(sqlite3_stmt*);
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_step = .CallFunc("sqlite3_step", sqlite3_stmt)
        End With
    Else
        sqlite3_step = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_reset(ByVal sqlite3_stmt As Long) As Long
    'int sqlite3_reset(sqlite3_stmt *pStmt);
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_reset = .CallFunc("sqlite3_reset", sqlite3_stmt)
        End With
    Else
        sqlite3_reset = SQLITE_ERROR
    End If
End Function

'const char *sqlite3_bind_parameter_name(sqlite3_stmt*, int)

Public Function sqlite3_bind_parameter_count(ByVal sqlite3_stmt As Long) As Long
    'int sqlite3_bind_parameter_count(sqlite3_stmt*);
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_bind_parameter_count = .CallFunc("sqlite3_bind_parameter_count", sqlite3_stmt)
        End With
    Else
        sqlite3_bind_parameter_count = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_bind_parameter_name(ByVal sqlite3_stmt As Long, ByVal Value As Long) As Long
    'const char *sqlite3_bind_parameter_name(sqlite3_stmt*, int)
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_bind_parameter_name = .CallFunc("sqlite3_bind_parameter_name", sqlite3_stmt, Value)
        End With
    Else
        sqlite3_bind_parameter_name = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_bind_parameter_index(ByVal sqlite3_stmt As Long, ByVal zName As String) As Long
    'int sqlite3_bind_parameter_index(sqlite3_stmt*, const char *zName);
    Dim bvStr() As Byte
    
    bvStr = StrConv(zName & vbNullChar, vbFromUnicode)
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_bind_parameter_index = .CallFunc("sqlite3_bind_parameter_index", sqlite3_stmt, VarPtr(bvStr(0)))
        End With
    Else
        sqlite3_bind_parameter_index = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_finalize(ByVal pStmt As Long) As Long
    'int sqlite3_finalize(sqlite3_stmt *pStmt);
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_finalize = .CallFunc("sqlite3_finalize", pStmt)
        End With
    Else
        sqlite3_finalize = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_column_bytes(ByVal sqlite3_stmt As Long, ByVal iCol As Long) As Long
    'int sqlite3_column_bytes(sqlite3_stmt*, int iCol);

    If m_bLoaded Then
        With m_cCDECL
            sqlite3_column_bytes = .CallFunc("sqlite3_column_bytes", sqlite3_stmt, iCol)
        End With
    End If
End Function

Public Function sqlite3_column_bytes16(ByVal sqlite3_stmt As Long, ByVal iCol As Long) As Long
    'int sqlite3_column_bytes16(sqlite3_stmt*, int iCol);

    If m_bLoaded Then
        With m_cCDECL
            sqlite3_column_bytes16 = .CallFunc("sqlite3_column_bytes16", sqlite3_stmt, iCol)
        End With
    End If
End Function

Public Function sqlite3_column_text(ByVal sqlite3_stmt As Long, ByVal iCol As Long) As String
    'const unsigned char *sqlite3_column_text(sqlite3_stmt*, int iCol);
    Dim tBLOB As DATA_BLOB
    
    If m_bLoaded Then
        With m_cCDECL
            tBLOB.pbData = .CallFunc("sqlite3_column_text", sqlite3_stmt, iCol)
            tBLOB.cbData = sqlite3_column_bytes(sqlite3_stmt, iCol)
            sqlite3_column_text = ReadBlobString(tBLOB)
        End With
    End If
End Function


Public Function sqlite3_column_text16(ByVal sqlite3_stmt As Long, ByVal iCol As Long) As String
    'const unsigned char *sqlite3_column_text(sqlite3_stmt*, int iCol);
    Dim tBLOB As DATA_BLOB
    
    If m_bLoaded Then
        With m_cCDECL
            tBLOB.pbData = .CallFunc("sqlite3_column_text16", sqlite3_stmt, iCol)
            tBLOB.cbData = sqlite3_column_bytes16(sqlite3_stmt, iCol)
            sqlite3_column_text16 = ReadBlobString(tBLOB)
            
            sqlite3_column_text16 = StrConv(sqlite3_column_text16, vbFromUnicode)
        End With
    End If
End Function

Public Function sqlite3_column_int(ByVal sqlite3_stmt As Long, ByVal iCol As Long) As Long
    'int sqlite3_column_int(sqlite3_stmt*, int iCol);
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_column_int = .CallFunc("sqlite3_column_int", sqlite3_stmt, iCol)
        End With
    End If
End Function

Public Function sqlite3_column_double(ByVal sqlite3_stmt As Long, ByVal iCol As Long) As Double
    'double sqlite3_column_double(sqlite3_stmt*, int iCol);
    Dim tBLOB As DATA_BLOB
    
    If m_bLoaded Then
        With m_cCDECL
            tBLOB.pbData = .CallFunc("sqlite3_column_text", sqlite3_stmt, iCol)
            tBLOB.cbData = sqlite3_column_bytes(sqlite3_stmt, iCol)
            sqlite3_column_double = CDbl(ReadBlobString(tBLOB))
        End With
    End If
End Function

Public Function sqlite3_column_type(ByVal sqlite3_stmt As Long, ByVal iCol As Long) As Long
    'int sqlite3_column_type(sqlite3_stmt*, int iCol);
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_column_type = .CallFunc("sqlite3_column_type", sqlite3_stmt, iCol)
        End With
    End If
End Function

Public Function sqlite3_column_count(ByVal sqlite3_stmt As Long) As Long
    'int sqlite3_column_count(sqlite3_stmt *pStmt);
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_column_count = .CallFunc("sqlite3_column_count", sqlite3_stmt)
        End With
    End If
End Function

Private Function sqlite3_column_blob(ByVal lStatement As Long, ByVal lCol As Long) As DATA_BLOB
    Dim tBLOB As DATA_BLOB
    
    'sqlite3_column_blob.pbData = CallFunc("sqlite3_column_blob", lStatement, lCol)
    'sqlite3_column_blob.cbData = sqlite3_column_bytes(lStatement, lCol)
End Function

Public Function sqlite3_column_name(ByVal sqlite3_stmt As Long, ByVal iCol As Long) As String
    'const char *sqlite3_column_name(sqlite3_stmt*, int N);
    Dim lRet As Long
    
    If m_bLoaded Then
        With m_cCDECL
            lRet = .CallFunc("sqlite3_column_name", sqlite3_stmt, iCol)
            sqlite3_column_name = StringFromPtr(lRet)
        End With
    End If
End Function

Public Function sqlite3_limit(ByVal sqlite3_stmt As Long, ByVal id As Long, ByVal NewVal As Long) As Long
    'int sqlite3_limit(sqlite3*, int id, int newVal);
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_limit = .CallFunc("sqlite3_limit", sqlite3_stmt, id, NewVal)
        End With
    Else
        sqlite3_limit = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_errcode(ByVal sqlite3 As Long) As Long
    'int sqlite3_errcode(sqlite3 *db);
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_errcode = .CallFunc("sqlite3_errcode", sqlite3)
        End With
    Else
        sqlite3_errcode = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_extended_errcode(ByVal sqlite3 As Long) As Long
    'int sqlite3_extended_errcode(sqlite3 *db);
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_extended_errcode = .CallFunc("sqlite3_extended_errcode", sqlite3)
        End With
    Else
        sqlite3_extended_errcode = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_errmsg(ByVal sqlite3 As Long) As String
    'const char *sqlite3_errmsg(sqlite3*);
    Dim lRet As Long
    
    If m_bLoaded Then
        With m_cCDECL
            lRet = .CallFunc("sqlite3_errmsg", sqlite3)
            sqlite3_errmsg = StringFromPtr(lRet)
        End With
    End If
End Function

Public Function sqlite3_libversion_number() As Long
    'int sqlite3_libversion_number(void);
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_libversion_number = .CallFunc("sqlite3_libversion_number")
        End With
    End If
End Function

Public Function sqlite3_libversion() As String
    'const char *sqlite3_libversion(void);
    Dim lRet As Long
    
    If m_bLoaded Then
        With m_cCDECL
            lRet = .CallFunc("sqlite3_libversion")
            sqlite3_libversion = StringFromPtr(lRet)
        End With
    End If
End Function

Public Function sqlite3_get_autocommit(ByVal sqlite3 As Long) As Long
    'int sqlite3_get_autocommit(sqlite3*);

    If m_bLoaded Then
        With m_cCDECL
            sqlite3_get_autocommit = .CallFunc("sqlite3_get_autocommit", sqlite3)
        End With
    End If
End Function

Public Function sqlite3_set_authorizer(ByVal sqlite3 As Long, ByVal xAuth As Long, ByVal pUserData As Long) As Long
    'int sqlite3_set_authorizer(
    '  sqlite3*,
    '  int (*xAuth)(void*,int,const char*,const char*,const char*,const char*),
    '  void *pUserData
    ');
    Dim lCallback As Long
    
    If m_bLoaded Then
        With m_cCDECL
            lCallback = .WrapCallback(xAuth, 6)
            sqlite3_set_authorizer = .CallFunc("sqlite3_set_authorizer", sqlite3, lCallback, pUserData)
        End With
    Else
        sqlite3_set_authorizer = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_sleep(ByVal Value As Long) As Long
    'int sqlite3_sleep(int);
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_sleep = .CallFunc("sqlite3_sleep", Value)
        End With
    Else
        sqlite3_sleep = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_soft_heap_limit(ByVal Value As Long) As Long
    'void sqlite3_soft_heap_limit(int);
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_soft_heap_limit = .CallFunc("sqlite3_soft_heap_limit", Value)
        End With
    Else
        sqlite3_soft_heap_limit = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_release_memory(ByVal Value As Long) As Long
    'int sqlite3_release_memory(int);
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_release_memory = .CallFunc("sqlite3_release_memory", Value)
        End With
    Else
        sqlite3_release_memory = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_malloc(ByVal Value As Long) As Long
    'void *sqlite3_malloc(int);
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_malloc = .CallFunc("sqlite3_malloc", Value)
        End With
    Else
        sqlite3_malloc = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_realloc(ByVal handle As Long, ByVal Value As Long) As Long
    'void *sqlite3_realloc(void*, int);
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_realloc = .CallFunc("sqlite3_realloc", handle, Value)
        End With
    Else
        sqlite3_realloc = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_free(ByVal handle As Long) As Long
    'void sqlite3_free(void*);
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_free = .CallFunc("sqlite3_free", handle)
        End With
    Else
        sqlite3_free = SQLITE_ERROR
    End If
End Function

Public Function sqlite3_randomness(ByVal n As Long, ByVal handle As Long) As Long
    'void sqlite3_randomness(int N, void *P);
    
    If m_bLoaded Then
        With m_cCDECL
            sqlite3_randomness = .CallFunc("sqlite3_randomness", n, handle)
        End With
    Else
        sqlite3_randomness = SQLITE_ERROR
    End If
End Function

Private Function ReadBlobString(ByRef tBLOB As DATA_BLOB) As String
    Dim b       As Byte
    Dim i       As Long
    
    If tBLOB.cbData = 0 Then Exit Function
    If tBLOB.pbData = 0 Then Exit Function

    For i = 0 To tBLOB.cbData - 1
        CopyMemory b, ByVal tBLOB.pbData + i, 1
        ReadBlobString = ReadBlobString & Chr$(b)
    Next
End Function

Public Function StringFromPtr(ByVal lAddr As Long) As String
    Dim b       As Byte
    
    If lAddr = 0 Then Exit Function
    Do
        CopyMemory b, ByVal lAddr, 1
        lAddr = lAddr + 1
        If b = 0 Then Exit Do
        StringFromPtr = StringFromPtr & Chr$(b)
    Loop
End Function

Public Function BytesFromPtr(ByVal lAddr As Long, ByVal lSize As Long) As Byte()
    Dim bvData() As Byte
    
    ReDim bvData(lSize - 1)
    CopyMemory bvData(0), ByVal lAddr, lSize
    BytesFromPtr = bvData
End Function

Private Function UnicodeToUTF8(ByVal sData As String) As String
    Dim bvData()    As Byte
    Dim lSize       As Long
    Dim lRet        As Long
    
    If LenB(sData) Then
        lSize = Len(sData) * 4
        ReDim bvData(lSize)
    
        lRet = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sData), _
           Len(sData), bvData(0), lSize + 1, vbNullString, 0)
    
        If lRet Then
            ReDim Preserve bvData(lRet - 1)
            UnicodeToUTF8 = StrConv(bvData, vbUnicode)
        End If
    End If

End Function

Private Function UnicodeToUTF16(ByVal sData As String) As String
    Dim bvData()    As Byte
    Dim lSize       As Long
    Dim lRet        As Long
    
    If LenB(sData) Then
    
        sData = UnicodeToUTF8(sData)
        
        lSize = Len(sData)
        ReDim bvData(lSize)
    
        lRet = MultiByteToWideChar(CP_UTF8, 0, StrPtr(sData), _
           Len(sData), bvData(0), lSize + 1)
    
        If lRet Then
            ReDim Preserve bvData(lRet - 1)
            UnicodeToUTF16 = StrConv(bvData, vbUnicode)
        End If
    End If

End Function

Public Function Deref(lAddr As Long) As Long
    Deref = lAddr
End Function

Private Function Canonicalize(sPath As String) As String
    Dim sBuff As String

    sBuff = Space$(260)

    If PathCanonicalize(sBuff, sPath) Then
        Canonicalize = Left$(sBuff, InStr(1, sBuff, vbNullChar) - 1)
    Else
        Canonicalize = sPath
    End If
End Function
