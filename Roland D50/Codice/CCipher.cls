VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCipher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************************
'  Copyright ©2004 Karl E. Peterson
'  All Rights Reserved, http://www.mvps.org/vb
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' Set aside storage for private member variables.
Private m_Mmf As CMapFile       ' Map all file i/o thru class.
Private m_Sa As CSafeArray      ' Map data into native array.
Private m_Data() As Byte        ' Array to serve data from.

' Default values.
Private Const defReadOnly As Boolean = False

' Notification events
Public Event FileMapped()
Public Event Progress(ByVal Percent As Long)


' ********************************************
'  Initialize / Terminate
' ********************************************
Private Sub Class_Initialize()
   ' Initialize private data.
   ReDim m_Data(0 To 0) As Byte
   Set m_Sa = New CSafeArray
   Set m_Mmf = New CMapFile
End Sub

Private Sub Class_Terminate()
   ' Clean up mapped array.
   Set m_Sa = Nothing
   ' Close open matrix file.
   Set m_Mmf = Nothing
End Sub

' ********************************************
'  Public Properties - r/w
' ********************************************
Public Property Let FileName(ByVal NewFile As String)
   ' Clear existing filemapping.
   m_Sa.ReleaseArray
   m_Mmf.UnMapFile
   
   ' Attempt to map new file to native VB array.
   If m_Mmf.MapFile(NewFile) Then
      Call MapDataToArray
   End If
End Property

Public Property Get FileName() As String
   ' Return filename from MMF.
   FileName = m_Mmf.MappedFile
End Property

' ********************************************
'  Public Properties - r/o
' ********************************************
Public Property Get FileSize() As Long
   ' Return filesize from MMF.
   FileSize = m_Mmf.MappedSize
End Property

' ********************************************
'  Public Methods
' ********************************************
Public Sub ApplyKey(ByVal Key As String, Optional Unicode As Boolean = False, Optional Alerts As Boolean = True)
   Dim b() As Byte
   Dim i As Long
   Dim pl As Long
   
   ' Transform key to byte array.
   If Unicode Then
      b = Key
   Else
      b = StrConv(Key, vbFromUnicode)
   End If
   
   ' Calculate length of password.
   pl = UBound(b) - LBound(b) + 1
   
   ' XOr successive bytes from key to file bytes.
   If Alerts Then
      RaiseEvent Progress(0)
      For i = LBound(m_Data) To UBound(m_Data)
         m_Data(i) = m_Data(i) Xor b(i Mod pl)
         If i Mod 10240 = 0 Then
            RaiseEvent Progress(i / UBound(m_Data) * 100)
         End If
      Next i
   
   Else
      For i = LBound(m_Data) To UBound(m_Data)
         m_Data(i) = m_Data(i) Xor b(i Mod pl)
      Next i
   End If
   
End Sub

' ********************************************
'  Private Methods
' ********************************************
Private Sub MapDataToArray()
   With m_Sa
      ' Attach current data to native array.
      .AttachToArray m_Data
      ' Set lowerbase of first (only) dimension to 0.
      .LowerBound(1) = 0
      ' Set number of elements to file size.
      .CountOfElements(1) = m_Mmf.MappedSize
      ' Assign actual data.
      .PointerToData = m_Mmf.BaseAddress
   End With
End Sub

