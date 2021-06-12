VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14475
   ClipControls    =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   14475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Hide"
      Height          =   375
      Left            =   12600
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   6588
      _Version        =   393216
      Rows            =   9
      Cols            =   9
      RowHeightMin    =   400
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   2
      ScrollBars      =   0
      AllowUserResizing=   3
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Visible = False
Form1.Visible = 1
End Sub

Private Sub Form_Load()
'Form2.Height = Grid.Height
For i = 1 To 8
    Form2.Grid.TextMatrix(i, 0) = i
    Form2.Grid.TextMatrix(0, i) = i
Next i
For colonna = 1 To 8
    Form2.Grid.ColWidth(colonna) = 1600
Next
End Sub

Private Sub Grid_Click()
current = ((Grid.Row - 1) * 8 + Grid.Col) - 1
If Form1.List1.ListCount = 0 Then Exit Sub
Form1.List2.ListIndex = current
ciao = Form1.List2.List(Form1.List2.ListIndex)
Call Form1.SendPatch
End Sub
Private Sub archive()

Dim Header

Const Esa = "0123456789ABCDEF"
Dim bytes() As Byte
Dim CharIndex As Integer
Dim ChCur As String
Header = "F0 41 00 14 12"
'RolandChar = " ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890-"
'DumpFile = "C:\prova.dmp"
freef = FreeFile
If DumpFile = "" Then Exit Sub
file_length = FileLen(DumpFile)
ReDim bytes(1 To file_length)
Open DumpFile For Binary Access Read As #freef
current = ((Grid.Row - 1) * 8 + Grid.Col) - 1
Offset = current * 448


strpatch = ""
sp = Chr$(32)
SyxByte = ""
List1.Clear
    For Patch = 0 To Offset - 1
           Get #freef, 1, bytes
           'StrPatch = StrPatch & Chr$(Format$(bytes(Patch)))
           'TransString = ""
           'ret = Len(StrPatch)
    Next Patch
    
    For Patch = Offset To Offset + 447
           Get #freef, 1, bytes
           StrByte = Format$(bytes(Patch + 1))
           SyxByte = SyxByte + Chr$(Val(StrByte))
           hexbyte = Hex$(StrByte)
           
           If Len(hexbyte) < 2 Then
           hexbyte = "0" + hexbyte
    End If
           HexBody = HexBody & hexbyte + sp
           
    Next Patch
    
'*** a questo punto è stata letta l'intera zona di memoria della patch
Text2.Text = HexBody
Label1.Caption = Str(Len(SyxByte)) + " Bytes"
Close #freef
sp = Chr$(32)
Chk = 0
ins = ""
newins = ""



'************* 1 istruzione sysex ***************
For X = 1 To 768 Step 3
    
    CurrentParam = Mid$(Text2.Text, X, 2)
    strLb = UCase(Right$(CurrentParam, 1))
    StrHb = UCase(Left$(CurrentParam, 1))
    ret = InStr(Esa, strLb)
    
    If ret = 0 Or strLb = "" Then
        errore = 1
        Exit For
    End If
    ret = InStr(Esa, StrHb)
    
    If ret = 0 Or StrHb = "" Then
        'Par(x).BackColor = &HFF
        errore = 1
        Exit For
    End If
    
    If Len(CurrentParam) < 2 Then
        CurrentParam = "0" + CurrentParam
    End If
    
    ins = ins + sp + UCase(CurrentParam)
    Chk = Chk + Val("&h" + CurrentParam)

Next X
If errore Then GoTo errore
'Chk = Chk + 2 ' aggiunto per l'aggiunta dei bytes di indirizzo dopo l'header
Chk = Chk And 127
Pivot = &H80 - Chk
While (Pivot < 0) Or (Pivot > 127)
    Chk = &H80 - (Chk - &H80)
    Pivot = Chk
Wend

Checksum = Hex$(Pivot)

If Pivot < 16 Then Checksum = "0" + Checksum

newins = Header + sp + "00 00 00" + ins + sp + Checksum + sp + "F7"

'List1.List(List1.ListIndex) = newins
List1.AddItem newins
'Debug.Print newins

sp = Chr$(32)
Chk = 0
ins = ""
newins = ""

' ************** 2 instruzione sysex ***************
For X = 769 To Len(Text2.Text) Step 3
  
    CurrentParam = Mid$(Text2.Text, X, 2)
    strLb = UCase(Right$(CurrentParam, 1))
    StrHb = UCase(Left$(CurrentParam, 1))
    ret = InStr(Esa, strLb)
    
    If ret = 0 Or strLb = "" Then
        errore = 1
        Exit For
    End If
    ret = InStr(Esa, StrHb)
    
    If ret = 0 Or StrHb = "" Then
       
        errore = 1
        Exit For
    End If
    
    If Len(CurrentParam) < 2 Then
        CurrentParam = "0" + CurrentParam
    End If
    ins = ins + sp + UCase(CurrentParam)
    Chk = Chk + Val("&h" + CurrentParam)
Next X
If errore Then GoTo errore
Chk = Chk + 2 ' aumentato per l'aggiunta dei bytes di indirizzo dopo l'header
Chk = Chk And 127
Pivot = &H80 - Chk

While (Pivot < 0) Or (Pivot > 127)
Chk = &H80 - (Chk - &H80)
Pivot = Chk
Wend

Checksum = Hex$(Pivot)

If Pivot < 16 Then Checksum = "0" + Checksum

newins = Header + sp + "00 02 00" + ins + sp + Checksum + sp + "F7"

'List1.List(List1.ListIndex) = newins
List1.AddItem newins

'Call Command1_Click
Call Form1.SendPatch
'Debug.Print newins
errore:

End Sub
