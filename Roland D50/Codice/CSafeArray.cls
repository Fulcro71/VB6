VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSafeArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************************
'  Written by Bill McCarthy, billmcc@iprimus.com.au, for VBPJ...
'  http://www.fawcette.com/Archives/premier/mgznarch/vbpj/2000/07jul00/bb0007/bb0007.asp
' *************************************************************************
Option Explicit

Public Enum SAFeatures
   FADF_AUTO = &H1            ' An array that is allocated on the stack.
   FADF_STATIC = &H2          ' An array that is statically allocated.
   FADF_EMBEDDED = &H4        ' An array that is embedded in a structure.
   FADF_FIXEDSIZE = &H10      ' An array that may not be resized or reallocated.
   FADF_RECORD = &H20         ' An array containing records. When set there will be a pointer to the IRecordinfo interface at negative offset 4 in the array descriptor.
   FADF_HAVEIID = &H40        ' An array that has an IID identifying interface. When set there will be a guid at negative offset 16 in the safearray descriptor. Flag is set only when FADF_DISPATCH or FADF_UNKNOWN is also set.
   FADF_HAVEVARTYPE = &H80    ' An array that has a VT type. When set there will be a VT tag at negative offset 4 in the array descriptor that specifies the element type.
   FADF_BSTR = &H100          ' An array of BSTRs.
   FADF_UNKNOWN = &H200       ' An array of IUnknown*.
   FADF_DISPATCH = &H400      ' An array of IDispatch*.
   FADF_VARIANT = &H800       ' An array of VARIANTs.
   FADF_RESERVED = &HF0E8     ' Bits reserved for future use.
End Enum

Private Type SAFEARRAYBOUND
   cElements As Long
   lLbound As Long
End Type

Private Type SAFEARRAY
   cDims As Integer           ' Count of dimensions in this array.
   fFeatures As Integer       ' Flags used by the SafeArray
   cbElements As Long         ' Size of an element of the array.
   cLocks As Long             ' Number of locks
   pvData As Long             ' Pointer to the data.
End Type

' Win32 API declarations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long

' Used to restore array's previous settings
Private m_pSA As Long
Private m_SAOld As SAFEARRAY
Private m_SAOldBounds() As SAFEARRAYBOUND

' Points directly to the safearray structure
Private m_Sa(0) As SAFEARRAY
Private m_SABounds() As SAFEARRAYBOUND

' Used to restore the pvData of the above to arrays
Private m_pvDataSA As Long
Private m_pvDataSABounds As Long

' ********************************************
'  Initialize / Terminate
' ********************************************
Private Sub Class_Terminate()
   ' Clean up!
   Me.ReleaseArray
End Sub

' ********************************************
'  Public Properties
' ********************************************
Public Property Get Dimensions() As Integer
   Dimensions = m_Sa(0).cDims
End Property

Public Property Let Dimensions(ByVal nDims As Integer)
   ' *** DANGEROUS ***
   If nDims <= m_SAOld.cDims Then
      m_Sa(0).cDims = nDims
   End If
End Property

Public Property Get Features() As SAFeatures
   Features = m_Sa(0).fFeatures
End Property

Public Property Let Features(ByVal iFeatures As SAFeatures)
   m_Sa(0).fFeatures = iFeatures
End Property

Public Property Get SizeofElements() As Long
   SizeofElements = m_Sa(0).cbElements
End Property

Public Property Let SizeofElements(ByVal lngElementsSize As Long)
   m_Sa(0).cbElements = lngElementsSize
End Property

Public Property Get Locks() As Long
   Locks = m_Sa(0).cLocks
End Property

Public Property Let Locks(ByVal nLocks As Long)
   m_Sa(0).cLocks = nLocks
End Property

Public Property Get PointerToData() As Long
   PointerToData = m_Sa(0).pvData
End Property

Public Property Let PointerToData(ByVal lngPtr As Long)
   m_Sa(0).pvData = lngPtr
End Property

Public Property Get LowerBound(Optional lngDim As Long = 1) As Long
   On Error Resume Next 'just in case it's out of bounds
   LowerBound = m_SABounds(lngDim).lLbound
End Property

Public Property Let LowerBound(Optional lngDim As Long = 1, lngLwrBound As Long)
   On Error Resume Next 'just in case it's out of bounds
   m_SABounds(lngDim).lLbound = lngLwrBound
End Property

Public Property Get CountOfElements(Optional lngDim As Long = 1) As Long
   On Error Resume Next 'just in case it's out of bounds
   CountOfElements = m_SABounds(lngDim).cElements
End Property

Public Property Let CountOfElements(Optional lngDim As Long = 1, lngCount As Long)
   On Error Resume Next 'just in case it's out of bounds
   m_SABounds(lngDim).cElements = lngCount
End Property

' ********************************************
'  Public Methods
' ********************************************
Public Sub AttachToArray(varArray)
   Dim ppSA As Long
   Dim pSA As Long
   Dim nDims As Integer
   Dim lngVarType As Long
   
   Call ReleaseArray
   
   CopyMemory lngVarType, ByVal VarPtr(varArray), 4&
   
   If lngVarType And &H2000& Then ' is an array
      'get pointer to safearray structure
      CopyMemory ppSA, ByVal VarPtr(varArray) + 8&, 4&
      
      'If varType And & H4000 = 0 Then variant array
      'and ppSA is actually pSA
      If lngVarType And &H4000& Then
         CopyMemory m_pSA, ByVal ppSA, 4&
      Else
         m_pSA = ppSA
      End If
   
   ElseIf lngVarType And &H3 Then
      'assume it's a pointer to an array, ppSA
      ppSA = varArray
      CopyMemory m_pSA, ByVal ppSA, 4&
   End If
   
   If m_pSA Then
      'get safearray structure
      CopyMemory nDims, ByVal m_pSA, 2&
      
      If nDims < 1 Then nDims = 1 'just in case
      
      ReDim m_SAOldBounds(nDims - 1)
      CopyMemory m_SAOld, ByVal m_pSA, 16&
      CopyMemory m_SAOldBounds(0), ByVal m_pSA + 16&, nDims * 8&
     
      ReDim m_SABounds(0)

      'point m_sa(0) to safearray structure
      CopyMemory pSA, ByVal VarPtrArray(m_Sa), 4&
      CopyMemory m_pvDataSA, ByVal pSA + 12&, 4&
      CopyMemory ByVal pSA + 12&, m_pSA, 4&
      
      'point m_SABounds() to safearray bounds
      CopyMemory pSA, ByVal VarPtrArray(m_SABounds), 4&
      CopyMemory m_pvDataSABounds, ByVal pSA + 12&, 4&
      CopyMemory ByVal pSA + 12&, m_pSA + 16&, 4&
      ' redim m_SABounds
      CopyMemory ByVal pSA + 16&, nDims, 4&
      CopyMemory ByVal pSA + 20&, 1&, 4&
      
      'lock the array
      m_Sa(0).cLocks = m_Sa(0).cLocks + 1&
   End If
End Sub

Public Sub ReleaseArray()
   Dim pSA As Long
   If m_pSA Then
      CopyMemory ByVal m_pSA, m_SAOld, 16&
      CopyMemory ByVal m_pSA + 16&, m_SAOldBounds(0), m_SAOld.cDims * 8&
      m_pSA = 0
      
      If m_pvDataSA Then
         CopyMemory pSA, ByVal VarPtrArray(m_Sa), 4
         CopyMemory ByVal pSA + 12, m_pvDataSA, 4
         m_pvDataSA = 0
      End If
      
      If m_pvDataSABounds Then
         CopyMemory pSA, ByVal VarPtrArray(m_SABounds), 4
         CopyMemory ByVal pSA + 12, m_pvDataSABounds, 4
         CopyMemory ByVal pSA + 16&, 1&, 4&
         CopyMemory ByVal pSA + 20&, 0&, 4&
         m_pvDataSABounds = 0
      End If
   End If
End Sub


