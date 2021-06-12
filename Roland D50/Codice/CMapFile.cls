VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMapFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************************
'  Copyright ©1998-2004 Karl E. Peterson
'  All Rights Reserved, http://www.mvps.org/vb
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' *************************************************
'  Win32 API Declarations
' *************************************************
Private Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FlushViewOfFile Lib "kernel32" (lpBaseAddress As Any, ByVal dwNumberOfBytesToFlush As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
' Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' *************************************************
' Used by the CreateFile API function
' *************************************************
Private Const NORMAL_FILE = 0
Private Const DELETE_ON_EXIT_FILE = 1
Private Const PAGEFILE_FILE = 2

Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Private Const FILE_READ_DATA = (&H1)                     '  file pipe
Private Const FILE_READ_ATTRIBUTES = (&H80)              '  all
Private Const FILE_READ_EA = (&H8)                       '  file directory
Private Const SYNCHRONIZE = &H100000
Private Const FILE_WRITE_DATA = (&H2)                    '  file pipe
Private Const FILE_WRITE_ATTRIBUTES = (&H100)            '  all
Private Const FILE_WRITE_EA = (&H10)                     '  file directory
Private Const FILE_APPEND_DATA = (&H4)                   '  file

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000

Private Const FILE_GENERIC_READ = _
   (STANDARD_RIGHTS_READ Or _
    FILE_READ_DATA Or _
    FILE_READ_ATTRIBUTES Or _
    FILE_READ_EA Or _
    SYNCHRONIZE)
Private Const FILE_GENERIC_WRITE = _
   (STANDARD_RIGHTS_WRITE Or _
    FILE_WRITE_DATA Or _
    FILE_WRITE_ATTRIBUTES Or _
    FILE_WRITE_EA Or _
    FILE_APPEND_DATA Or _
    SYNCHRONIZE)

Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Const CREATE_NEW = 1
Private Const CREATE_ALWAYS = 2
Private Const OPEN_EXISTING = 3
Private Const OPEN_ALWAYS = 4
Private Const TRUNCATE_EXISTING = 5

Private Const FILE_FLAG_WRITE_THROUGH = &H80000000
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const FILE_FLAG_NO_BUFFERING = &H20000000
Private Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Private Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Private Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Private Const FILE_FLAG_POSIX_SEMANTICS = &H1000000

' *************************************************
' Used by the OpenFileMapping API function
' *************************************************
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SECTION_QUERY = &H1
Private Const SECTION_MAP_WRITE = &H2
Private Const SECTION_MAP_READ = &H4
Private Const SECTION_MAP_EXECUTE = &H8
Private Const SECTION_EXTEND_SIZE = &H10
Private Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE

Private Const FILE_MAP_COPY = SECTION_QUERY
Private Const FILE_MAP_WRITE = SECTION_MAP_WRITE
Private Const FILE_MAP_READ = SECTION_MAP_READ
Private Const FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS

Private Const PAGE_READONLY = &H2
Private Const PAGE_READWRITE = &H4
Private Const PAGE_WRITECOPY = &H8

' *************************************************-------
' Eventlog Status Codes
' *************************************************-------
Private Const INVALID_HANDLE_VALUE = -1&
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_INVALID_FUNCTION = 1&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_TOO_MANY_OPEN_FILES = 4&
Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_INVALID_HANDLE = 6&
Private Const ERROR_ARENA_TRASHED = 7&
Private Const ERROR_NOT_ENOUGH_MEMORY = 8&
Private Const ERROR_INVALID_BLOCK = 9&
Private Const ERROR_BAD_ENVIRONMENT = 10&
Private Const ERROR_INVALID_ACCESS = 12&
Private Const ERROR_INVALID_DATA = 13&
Private Const ERROR_BAD_FORMAT = 11&
Private Const ERROR_OUTOFMEMORY = 14&
Private Const ERROR_HANDLE_EOF = 38&
Private Const ERROR_INVALID_PARAMETER = 87&
Private Const ERROR_INSUFFICIENT_BUFFER = 122&
Private Const ERROR_INVALID_NAME = 123&
Private Const ERROR_ALREADY_EXISTS = 183&
Private Const ERROR_MORE_DATA = 234&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const ERROR_PARTIAL_COPY = 299&
Private Const ERROR_INVALID_ADDRESS = 487&
Private Const ERROR_NOACCESS = 998&
Private Const ERROR_CANTOPEN = 1011&
Private Const ERROR_CANTREAD = 1012&
Private Const ERROR_CANTWRITE = 1013&
Private Const ERROR_MAPPED_ALIGNMENT = 1132&

' *************************************************
' Public -- keep track of our own error info
' *************************************************
Private Const ERR_FAILED_CLOSE_MEMORY_MAPPING = 1001
Private Const ERR_FAILED_CLOSE_MM_FILE = 1002
Private Const ERR_FAILED_UNMAP_VIEW_FILE = 1003
Private Const ERR_PASSED_NULL_TEXT = 1101
Private Const ERR_WRITE_TO_FILE = 1102
Private Const ERR_READ_FROM_FILE = 1103
Private Const ERR_BYTES_TO_MAP_TOO_BIG = 1201
Private Const ERR_INVALID_MM_FILENAME = 1202
Private Const ERR_WRONG_FILE_TYPE = 1203
Private Const ERR_INVALID_TMP_FILE = 1204
Private Const ERR_CREATE_MM_FILE = 1205
Private Const ERR_CREATE_MEM_MAPPING = 1206
Private Const ERR_CREATE_FILE_VIEW = 1207

' *************************************************
' Member variables
' *************************************************
Private m_hFile As Long
Private m_BaseAddr As Long
Private m_FileName As String
Private m_FileSize As Long
Private m_Paged As Boolean

' *************************************************
'  Initialization/Termination
' *************************************************
Private Sub Class_Initialize()
   '
End Sub

Private Sub Class_Terminate()
   ' close open file if need be...
   If m_hFile Then Call UnMapFile
End Sub

' *************************************************
'  Public Properties (Read/Write)
' *************************************************


' *************************************************
'  Public Properties (Read-Only)
' *************************************************
Public Property Get BaseAddress() As Long
   BaseAddress = m_BaseAddr
End Property

Public Property Get hFile() As Long
   hFile = m_hFile
End Property

Public Property Get MappedFile() As String
   MappedFile = m_FileName
End Property

Public Property Get MappedSize() As Long
   MappedSize = m_FileSize
End Property

Public Property Get Paged() As Boolean
   Paged = m_Paged
End Property

' *************************************************
'  Public Access Methods
' *************************************************
Public Function GetDbl(ByVal Offset As Long) As Double
   If m_BaseAddr Then
      CopyMemory GetDbl, ByVal (m_BaseAddr + Offset), 8
   End If
End Function

Public Function GetLng(ByVal Offset As Long) As Long
   If m_BaseAddr Then
      CopyMemory GetLng, ByVal (m_BaseAddr + Offset), 4
   End If
End Function

Public Sub GetRng(ByVal lpDestination As Long, ByVal Offset As Long, ByVal Length As Long)
   If m_BaseAddr Then
      CopyMemory ByVal lpDestination, ByVal (m_BaseAddr + Offset), Length
   End If
End Sub

Public Sub PutDbl(ByVal PutVal As Double, ByVal Offset As Long)
   If m_BaseAddr Then
      CopyMemory ByVal (m_BaseAddr + Offset), PutVal, 8
   End If
End Sub

Public Function PutLng(ByVal PutVal As Long, ByVal Offset As Long)
   If m_BaseAddr Then
      CopyMemory ByVal (m_BaseAddr + Offset), PutVal, 4
   End If
End Function

Public Sub PutRng(ByVal lpSource As Long, ByVal Offset As Long, ByVal Length As Long)
   If m_BaseAddr Then
      CopyMemory ByVal (m_BaseAddr + Offset), ByVal lpSource, Length
   End If
End Sub

' *************************************************
'  Public Mapping Methods
' *************************************************
Public Function MapFile(ByVal FileName As String) As Boolean
   Const Disposition As Long = OPEN_EXISTING
   MapFile = OpenAndMap(FileName, 0, Disposition)
End Function

Public Function MapFileNew(ByVal FileName As String, ByVal Size As Long) As Boolean
   Const Disposition As Long = CREATE_ALWAYS
   MapFileNew = OpenAndMap(FileName, Size, Disposition)
End Function

Public Function MapFileTemp(ByVal Size As Long) As Boolean
   MapFileTemp = OpenAndMapPaged(Size)
End Function

Public Function UnMapFile() As Boolean
   ' unmap the current view...
   If UnmapViewOfFile(m_BaseAddr) <> 0 Then
      Call CloseHandle(m_hFile)
      m_hFile = 0
      m_BaseAddr = 0
      m_FileName = ""
      m_FileSize = 0
      m_Paged = False
      UnMapFile = True
   End If
End Function

' *************************************************
'  Private Methods
' *************************************************
Private Function OpenAndMap(ByVal FileName As String, ByVal Size As Long, ByVal Disposition As Long) As Boolean
   Dim hFile As Long
   Dim nRet As Long

   'Const CreateFlags As Long = FILE_GENERIC_WRITE Or FILE_GENERIC_READ
   Const CreateFlags As Long = GENERIC_WRITE Or GENERIC_READ
   Const ShareFlags As Long = FILE_SHARE_READ Or FILE_SHARE_WRITE

   ' make sure we don't already have a file open...
   If m_hFile Then Call UnMapFile

   ' open existing file, or bail...
   If Len(FileName) Then
      hFile = CreateFile(FileName, CreateFlags, ShareFlags, ByVal 0&, Disposition, FILE_FLAG_RANDOM_ACCESS, 0&)
   End If
   If hFile <> INVALID_HANDLE_VALUE Then

      ' create file mapping, requesting existing size as maximum...
      m_hFile = CreateFileMapping(hFile, ByVal 0&, PAGE_READWRITE, 0&, Size, vbNullString)
      If m_hFile Then

         ' map the entire file into this process's address space...
         m_BaseAddr = MapViewOfFile(m_hFile, FILE_MAP_ALL_ACCESS, 0&, 0&, 0&)

         ' if we have an address, we're off to the races!
         If m_BaseAddr <> 0 Then
            m_FileName = FileName
            m_FileSize = GetFileSize(hFile, ByVal 0&)
            OpenAndMap = True
         End If
      End If

      ' we don't need original file handle anymore...
      Call CloseHandle(hFile)
   End If

   ' This type of file is always disk-based.
   m_Paged = False
End Function

Private Function OpenAndMapPaged(ByVal Size As Long) As Boolean
   Dim nRet As Long
   Const hFile As Long = &HFFFFFFFF

   ' make sure we don't already have a file open...
   If m_hFile Then Call UnMapFile

   ' create file mapping, backed by the pagefile...
   m_hFile = CreateFileMapping(hFile, ByVal 0&, PAGE_READWRITE, 0&, Size, vbNullString)
   If m_hFile Then

      ' map the entire file into this process's address space...
      m_BaseAddr = MapViewOfFile(m_hFile, FILE_MAP_ALL_ACCESS, 0&, 0&, 0&)

      ' if we have an address, we're off to the races!
      If m_BaseAddr <> 0 Then
         m_FileName = "[Paged]"
         m_FileSize = Size
         m_Paged = True
         OpenAndMapPaged = True
      Else
         m_Paged = False
      End If
   End If
End Function


