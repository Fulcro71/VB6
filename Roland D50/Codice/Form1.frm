VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Credits"
      Height          =   4695
      Left            =   6480
      TabIndex        =   27
      Top             =   3240
      Width           =   10695
      Begin VB.PictureBox Picture1 
         Height          =   2535
         Left            =   360
         Picture         =   "Form1.frx":08CA
         ScaleHeight     =   2475
         ScaleWidth      =   6915
         TabIndex        =   28
         Top             =   960
         Width           =   6975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fulcro 2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   30
         Top             =   3600
         Width           =   1260
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1785
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   5760
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   5730
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "SysEx Editing && Validating"
      Height          =   4695
      Left            =   480
      TabIndex        =   8
      Top             =   2160
      Width           =   10695
      Begin VB.CommandButton Command3 
         Caption         =   "Load File"
         Height          =   375
         Left            =   8520
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Recheck"
         Height          =   375
         Left            =   8520
         TabIndex        =   14
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "Form1.frx":6C79
         Top             =   2280
         Width           =   8295
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         IntegralHeight  =   0   'False
         ItemData        =   "Form1.frx":6C7F
         Left            =   120
         List            =   "Form1.frx":6C81
         TabIndex        =   12
         Top             =   3600
         Width           =   8295
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Send"
         Height          =   375
         Left            =   8520
         TabIndex        =   11
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Clear"
         Height          =   375
         Left            =   8520
         TabIndex        =   10
         Top             =   960
         Width           =   1695
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         IntegralHeight  =   0   'False
         ItemData        =   "Form1.frx":6C83
         Left            =   120
         List            =   "Form1.frx":6C85
         TabIndex        =   9
         Top             =   360
         Width           =   8295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "File managing"
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   10695
      Begin VB.CommandButton Command14 
         Caption         =   "Patch Table"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Scan MIDI"
         Height          =   375
         Left            =   9000
         TabIndex        =   21
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   6000
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   8040
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Test Note"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Save List"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1575
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3900
         Left            =   2040
         TabIndex        =   5
         Top             =   600
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Send Patch"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Load Module"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   6000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1920
         Width           =   4455
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1440
         Top             =   2520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label5 
         Caption         =   "Selected Patch Dump:"
         Height          =   255
         Left            =   6960
         TabIndex        =   26
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Selected MIDI Interface:"
         Height          =   255
         Left            =   6120
         TabIndex        =   25
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Patch List:"
         Height          =   195
         Left            =   2160
         TabIndex        =   24
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Delay (ms):"
         Height          =   195
         Left            =   7200
         TabIndex        =   22
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   6960
         TabIndex        =   6
         Top             =   3720
         Width           =   2175
      End
   End
   Begin ComctlLib.TabStrip Tab1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   9128
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sender"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Advanced"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "About"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' 32 bit
'
'Option Base 0


Private Const LVM_FIRST As Long = &H1000
Private Const LVM_HITTEST As Long = (LVM_FIRST + 18)
Private Const LVM_SUBITEMHITTEST As Long = (LVM_FIRST + 57)
Private Const LVHT_ONITEMICON As Long = &H2
Private Const LVHT_ONITEMLABEL As Long = &H4
Private Const LVHT_ONITEMSTATEICON As Long = &H8
Private Const LVHT_ONITEM As Long = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)
Private Type POINTAPI
  X As Long
  Y As Long
End Type
Private Type LVHITTESTINFO
   pt As POINTAPI
   flags As Long
   iItem As Long
   iSubItem As Long
End Type


Private Const LB_SETHORIZONTALEXTENT = &H194

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer



Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Private Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function midiOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Private Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function midiOutGetErrorText Lib "winmm.dll" Alias "midiOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Private Declare Function MIDIOutOpen Lib "winmm.dll" Alias "midiOutOpen" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutPrepareHeader Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Private Declare Function midiOutUnprepareHeader Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Private Declare Function midiOutLongMsg Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long


Dim OutDevCount As Integer
Dim outSYX
Dim InsString As String
Private WithEvents crypt As CCipher
Attribute crypt.VB_VarHelpID = -1
Dim DumpFile As String
Dim Syxfile As String


Private Sub FillPatchList(DmpFile As String)
Dim bytes() As Byte
Dim CharIndex, R, C As Integer
Dim ChCur As String
RolandChar = " ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890-"
'DumpFile = "C:\prova.dmp"
freef = FreeFile
file_length = FileLen(DmpFile)
ReDim bytes(1 To file_length)
Open DmpFile For Binary Access Read As #freef
'For Punt = 1 To 64
strpatch = ""
'ProgressBar1.Value = 0

'ProgressBar1.Max = 28672

    For Patch = 1 To 28672
           Get #freef, 1, bytes
          strpatch = strpatch & Chr$(Format$(bytes(Patch)))
           TransString = ""
           ret = Len(strpatch)
           ProgressBar1.Value = ProgressBar1.Value + 1
    Next Patch
    
'*** a questo punto è stata letta l'intera zona di memoria delle patches (28672=448bytes per patch * 64 patches)
    R = 1
    C = 1
    
    For Punt = 0 To 63
        TransString = ""
        
    For intName = 1 To 18
        
            ChCur = Mid$(strpatch, (384 + 448 * Punt) + intName, 1)
            'CharIndex = InStr(ChCur, RolandChar)
            CharIndex = Asc(ChCur)
            TransChar = Mid$(RolandChar, CharIndex + 1, 1)
            TransString = TransString + TransChar
            
        Next intName
            Form2.Grid.Row = R                '*****************************
            Form2.Grid.Col = C                '* inserimento nella Grid
            Form2.Grid.Text = TransString     '*****************************
            
            List2.AddItem Trim(Trim(Str(R)) + Trim(Str(C)) + " " & TransString)
            If C < 8 Then
                        C = C + 1
                     Else
                        C = 1
                        R = R + 1
                     End If
            
    Next Punt



'Next Punt
    
'PatchName=
    
'Open SysFile For Binary As #freef


'   Parse (Systring)

Close #freef
'For i = 1 To file_length
        
        'strPatch = Format$(bytes(i))
        'If Len(HexByte) < 2 Then HexByte = "0" + HexByte
        'txt = txt + HexByte + Chr$(32)
        'If UCase(HexByte) = "F7" Then
        '                            List1.AddItem txt
        '                            txt = ""
        '                          End If
 '   Next i
 ProgressBar1.Visible = 0
End Sub



Private Sub Combo1_Click()
StatusBar1.Panels(3).Text = "Current MIDI Interface: " & Combo1.Text
End Sub

Private Sub Command1_Click()
Call SendPatch
End Sub

Public Sub SendPatch()
For i = 0 To List1.ListCount - 1

StrCurIns = List1.List(i)
ins = Chr$(Val("&h" + Trim(Left$(StrCurIns, 2))))
For Cur = 3 To (Len(StrCurIns) - 2) Step 3
        ins = ins + Chr$(Val("&h" + Trim(Mid$(StrCurIns, Cur, 3))))
    Next Cur
syx = ins
        outSYX = syx
        Call SYXSend
Next i
End Sub

Private Sub Command10_Click()
Text3.Text = ""
List3.Clear
List1.Clear
End Sub

Private Sub Command11_Click()


'****************************************************************
'*** Caricamento di file .syx e conversione in -dmp per i nomi
'****************************************************************
Dim file_name As String
    Dim file_length As Long
    Dim fnum As Integer
    Dim bytes() As Byte
    Dim txt As String
    Dim i As Long
Dim test As CCipher
'Dim Sysfile As String
Dim DmpFile As String
ProgressBar1.Value = 0
CommonDialog1.Filter = "Sysex File       (*.syx)|*.syx|D50 Sound Banks  (*.d50)|*.d50"
CommonDialog1.CancelError = False
CommonDialog1.ShowOpen
Syxfile = CommonDialog1.FileName

Set test = New CCipher

test.FileName = Syxfile
' a questo punto l'array con i byte del file è stato creato.
If Syxfile = "" Then Exit Sub

List1.Clear
List2.Clear
List3.Clear
Text3.Text = ""

Dim txtName As String
Dim ret As Integer
txtName = String(255, 0)
ret = GetFileTitle(Syxfile, txtName, Len(txtName))
txtName = Left$(txtName, InStr(1, txtName, Chr$(0)) - 1)
StatusBar1.Panels.Item(2).Text = "Current File: " & txtName

Form1.MousePointer = 11
DoEvents

file_length = FileLen(Syxfile)
If file_length < 36048 Then
    Form1.MousePointer = 0
    ret = MsgBox("File too short!", vbCritical + vbOKOnly, "Invalid File")
    Exit Sub
End If

freef = FreeFile
ReDim bytes(1 To file_length)
ProgressBar1.Visible = 1

Open Syxfile For Binary As #freef

   Get #freef, 1, bytes
'   Parse (Systring)

Close #freef
ProgressBar1.Max = 28672 + file_length
For i = 1 To file_length
        'txt = txt & Hex$(Format$("00" & (Format$(bytes(i)), "!&&")) & Chr$(32)
        hexbyte = Hex$(Format$(bytes(i)))
        If Len(hexbyte) < 2 Then hexbyte = "0" + hexbyte
        txt = txt + hexbyte + Chr$(32)
        If (i = 1) And (hexbyte <> "F0") Then
        Form1.MousePointer = 0
         ret = MsgBox("Corrupted file!", vbCritical + vbOKOnly, "Invalid File")
         
         Exit Sub
        End If
        If UCase(hexbyte) = "F7" Then
                                    List1.AddItem txt
                                    txt = ""
                                  End If
    ProgressBar1.Value = ProgressBar1.Value + 1
    Next i
    
    
Dim ins As String


'*** Conversione in .dmp

'If SysFile = "" Then Exit Sub

DmpFile = Left$(Syxfile, Len(Syxfile) - 4) & ".dmp"
Open DmpFile For Binary Access Write As #freef
'
 '  Get #freef, 1, bytes
'   Parse (Systring)

'Close #freef

For C = 0 To List1.ListCount - 1

FullIns = List1.List(C)
Dump = Mid$(FullIns, 25, Len(FullIns) - 31)
'StrCurIns = UCase(dump)
If Not (Dump = "") Then
ins = Chr$(Val("&h" + Trim(Left$(Dump, 2))))
For Cur = 3 To (Len(Dump) - 2) Step 3
        ins = ins + Chr$(Val("&h" + Trim(Mid$(Dump, Cur, 3))))
    Next Cur
End If

'Debug.Print Dump
Put #freef, , ins

Next C
Close #freef

        
Call FillPatchList(DmpFile)
DumpFile = DmpFile 'Dumpfile=variabile globale usata in altri moduli
Form1.MousePointer = 0
End Sub

Private Sub Command12_Click()

End Sub

Private Sub Command13_Click()
Dim ret As Integer
Dim txtName As String
Dim intName As Integer
txtName = String(255, 0)
'BankName = txtName
If Syxfile = "" Then Exit Sub
ret = GetFileTitle(Syxfile, txtName, Len(txtName))
txtName = Left$(txtName, InStr(1, txtName, Chr$(0)) - 1)
BankName = txtName
'MsgBox txtName
txtName = Left$(txtName, Len(txtName) - 3) + "txt"
'MsgBox txtName
freef = FreeFile
Dim listfile As String
listfile = txtName
CommonDialog1.FileName = listfile
CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
CommonDialog1.CancelError = False
CommonDialog1.flags = 2

'On Error Resume Next

replace:
CommonDialog1.ShowSave
'MsgBox err.Number
'If (err.Number = 32755) Then  ' pulsante Annulla
 '   Exit Sub
'End If
    listfile = CommonDialog1.FileName
If listfile = "" Then Exit Sub

'Dim fso
'Set fso = CreateObject("Scripting.FileSystemObject")
'If fso.FileExists(listfile) Then
 '   ret = MsgBox("File existing: overwrite?", vbOKCancel + vbInformation, "Confirm overwrite")
    'MsgBox ret
'End If
 '  If ret = 2 Then GoTo replace
   ' DoEvents




'SysFile = "C:\prova.dmp"
'ReDim bytes(1 To file_length)
'Open listfile For Binary Access Write As #freef

Open listfile For Output As #freef

 '  Get #freef, 1, bytes
'   Parse (Systring)

'Close #freef

Print #freef, "Roland D50 Sound Bank"
Print #freef, "File Name: " & BankName
Print #freef, ""
Print #freef, "Patch list:"
Print #freef, ""
For C = 0 To List2.ListCount - 1

PatchName = List2.List(C)
Print #freef, PatchName

Next C
Close #freef
Exit Sub
errmanage:

  'Else
    'MsgBox "[" & err.Number & "]: " & err.Description, vbCritical
   ' ret = MsgBox("File existing: overwrite?", vbOKCancel + vbInformation, "Confirm overwrite")
    'MsgBox ret
   ' If ret = 1 Then GoTo replace
   ' DoEvents
  'End If
End Sub

Private Sub Command14_Click()
'Frame2.Visible = 1
'Frame2.ZOrder
Form2.Caption = "Current MIDI Interface: " & Combo1.Text
Form2.Show 1

End Sub

Private Sub Command15_Click()
Frame2.Visible = 0
End Sub

Private Sub Command2_Click()
 Dim rc As Integer, tm
      'MyMIDI = Str$(VBSYXMID.Text.Text)
      MyMIDI = Str$(Combo1.ListIndex)
    If MyMIDI <> "N" Then
        mDev = (Val(MyMIDI)) - 1
        rc = MIDIOutOpen(hMidi, mDev, 0&, 0&, 0&)
        If rc <> 0 Then
            Call MidiErr("Open", rc)
            Exit Sub
        End If
        rc = midiOutShortMsg(hMidi, &H7F3C90) ' middle c note on velocity 127
        tm = Timer
        
        While tm > Timer - 1    'attesa di un secondo
        Wend
        
        rc = midiOutShortMsg(hMidi, &H7F3C80) ' middle c note off velocity 127
        rc = midiOutClose(hMidi)
        If rc <> 0 Then
            Call MidiErr("Close", rc)
        End If
    End If
End Sub

Private Sub Command3_Click()
Dim file_name As String
    Dim file_length As Long
    Dim fnum As Integer
    Dim bytes() As Byte
    Dim txt As String
    Dim i As Long


'**************************
Dim SyString As String
Dim SySFil As String
Dim test As CCipher

InsStrng = ""
CommonDialog1.Filter = "Sysex File (*.syx)|*.syx"
CommonDialog1.CancelError = False
CommonDialog1.ShowOpen

Sysfile = CommonDialog1.FileName
Set test = New CCipher

test.FileName = Sysfile
' a questo punto l'array con i byte del file è stato creato.
If Sysfile = "" Then Exit Sub
List1.Clear

'ret = GetShortPathName(SysFile, SySFil, 255)
'Me.Caption = Sysfile
file_length = FileLen(Sysfile)
freef = FreeFile
ReDim bytes(1 To file_length)
Open Sysfile For Binary As #freef

   Get #freef, 1, bytes
'   Parse (Systring)

Close #freef
For i = 1 To file_length
        'txt = txt & Hex$(Format$("00" & (Format$(bytes(i)), "!&&")) & Chr$(32)
        hexbyte = Hex$(Format$(bytes(i)))
        If Len(hexbyte) < 2 Then hexbyte = "0" + hexbyte
        txt = txt + hexbyte + Chr$(32)
        If UCase(hexbyte) = "F7" Then
                                    List1.AddItem txt
                                    txt = ""
                                  End If
    Next i
    
'Parse (txt)
'Command1.Enabled = True

End Sub

Private Sub Command4_Click()
List1.Clear

End Sub

Private Sub Command5_Click()
Call MIDI_Scan

End Sub

Private Sub Command6_Click()
'ModIns = Str(List1.List(List1.ListIndex))
'Checksum (ModIns)
Call ReCheck(Text3.Text)
End Sub

Private Sub Eliminare_Click()
freef = FreeFile
Dim ins As String

CommonDialog1.Filter = "Dump Files (*.dmp)|*.dmp"
CommonDialog1.CancelError = False
CommonDialog1.ShowSave
Sysfile = CommonDialog1.FileName
If Sysfile = "" Then Exit Sub

'SysFile = "C:\prova.dmp"
'ReDim bytes(1 To file_length)
Open Sysfile For Binary Access Write As #freef
'
 '  Get #freef, 1, bytes
'   Parse (Systring)

'Close #freef

For C = 0 To List1.ListCount - 1

FullIns = List1.List(C)
Dump = Mid$(FullIns, 25, Len(FullIns) - 31)
'StrCurIns = UCase(dump)
If Not (Dump = "") Then
ins = Chr$(Val("&h" + Trim(Left$(Dump, 2))))
For Cur = 3 To (Len(Dump) - 2) Step 3
        ins = ins + Chr$(Val("&h" + Trim(Mid$(Dump, Cur, 3))))
    Next Cur
End If

'Debug.Print Dump
Put #freef, , ins

Next C
Close #freef

    
    
End Sub

Private Sub Command8_Click()
'CommonDialog1.Filter = "Dump Files (*.dmp)|*.dmp"
'CommonDialog1.CancelError = False
'CommonDialog1.ShowOpen
'DumpFile = CommonDialog1.FileName
'If DumpFile = "" Then Exit Sub
'List2.Clear
'Call FillPatchList(DumpFile)
End Sub

Private Sub Command9_Click()
For i = 0 To List3.ListCount - 1

StrCurIns = List3.List(i)
ins = Chr$(Val("&h" + Trim(Left$(StrCurIns, 2))))
For Cur = 3 To (Len(StrCurIns) - 2) Step 3
        ins = ins + Chr$(Val("&h" + Trim(Mid$(StrCurIns, Cur, 3))))
    Next Cur
syx = ins
        outSYX = syx
        Call SYXSend
Next i
End Sub

Private Sub Form_Load()
Apptitle = "Roland D50-D550 Bank Editor"
Me.Caption = Apptitle
Label6.Caption = Apptitle

Text1.Text = 100
Text3.Text = ""

Frame1.Visible = 1
Frame2.Left = 120
Frame3.Left = 120
'Frame4.Left = 120
Frame2.Top = 480
Frame3.Top = 480
'Frame4.Top = 480

Frame1.ZOrder
ProgressBar1.Visible = 0
'Grid.Row = 1
'Grid.Col = 1
'Grid.Text = "prova"
Form2.Grid.BandIndent(0) = 1
PortForm = Form1.Width / 3
StatusBar1.Panels.Item(1).Width = PortForm
StatusBar1.Panels.Item(2).Width = PortForm
StatusBar1.Panels.Item(3).Width = PortForm
ProgressBar1.Width = PortForm



SendMessage List1.hwnd, LB_SETHORIZONTALEXTENT, 6500, 0&
SendMessage List3.hwnd, LB_SETHORIZONTALEXTENT, 6500, 0&




'For colonna = 1 To 8
'    Form2.Grid.ColWidth(colonna) = 1800
'Next
Call MIDI_Scan
End Sub
Private Sub MIDI_Scan()
Dim OutCaps As MIDIOUTCAPS
    Combo1.Clear
    Combo1.AddItem "Not Enabled"
    
    OutDevCount = midiOutGetNumDevs()
    For zz = 0 To OutDevCount - 1            ' Midi Mapper = -1
        vntRet = midiOutGetDevCaps(zz, OutCaps, Len(OutCaps))
        If vntRet <> 0 Then
            MsgBox "midiOutGetDevCaps Error: " & vntRet
            Exit For
        End If
        Combo1.AddItem OutCaps.szPname
        Next zz
         
         Combo1.ListIndex = 0
'Command1.Enabled = False
End Sub


Private Sub SYXSend()
    Dim rc As Integer
    MyMIDI = Str$(Combo1.ListIndex)
    If MyMIDI <> "N" Then
        mDev = (Val(MyMIDI)) - 1
        rc = MIDIOutOpen(hMidi, mDev, 0&, 0&, 0&)
        If rc <> 0 Then
            Call MidiErr("Open", rc)
            Exit Sub
        End If
        LongMidiMessage (outSYX)
        rc = midiOutClose(hMidi)
        If rc <> 0 Then
            Call MidiErr("Close", rc)
        End If
    End If
End Sub

Private Sub MidiErr(mOpt As String, rc As Integer)
    Dim msgText As String * 132
    
    vntRet = midiOutGetErrorText(rc, msgText, 128)
    MsgBox "Operation: " & mOpt & Chr(13) & Chr(10) & msgText
End Sub
Private Sub Parse(sys As String)
Dim Esa
Dim IntChr As Byte
Esa = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F")
'InsStrng = ""
Dim newline As Boolean
newline = False

For Length = 1 To Len(sys)
    curChar$ = Mid$(sys, Length, 1)
    'InsString = InsString + ("&h" + Val(curChar$))
    IntChr = AscB(curChar$)
    HChr = Int(IntChr / 16)
    LChr = IntChr - HChr * 16
    HexHChr = Esa(HChr)
    HexLChr = Esa(LChr)
    hexbyte = HexHChr + HexLChr + Chr$(32)
    If Trim(UCase(hexbyte) = "F7") Then newline = True
    InsString = InsString + hexbyte
    If newline Then
            'IntString = IntString + TxtSys
            List1.AddItem Trim(InsString)
            InsString = ""
            End If
            'IntString = IntString + TxtSys
    Next Length
    'IntString = IntString + TxtSys
    
    
End Sub

Private Sub Grid_Click()
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
Call SendPatch
'Debug.Print newins
errore:
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_Click()
FullIns = List1.List(List1.ListIndex)
IntLen = Len(FullIns)
FullIns = Mid$(FullIns, 16, IntLen - 21)
Text3.Text = FullIns
End Sub
Private Sub LongMidiMessage(InString As String)
    Dim mHdr As MIDIHDR
    Dim rc As Integer
    Dim Length As Integer
    Dim Checks As Integer
    Dim midistring(267) As Byte ' Make sure this is big enough!
    'Sleep (1000)
    ' This is wrong - you cannot use strings in MIDIHDR
    'Length = Len(InString)
    'mHdr.lpData = InString
    '
    ' Let's do this instead
    Length = 255
    For rc = 1 To Len(InString)
        midistring(rc - 1) = Asc(Mid$(InString, rc, 1))
    Next rc
    midistring(rc - 1) = 0 ' end of string - just in case :-)
    mHdr.lpData = VarPtr(midistring(0)) ' Undocumented feature!
    '
    mHdr.dwBufferLength = Length
    mHdr.dwBytesRecorded = 0 ' Was Length - only used for MIDI in
    mHdr.dwUser = 0
    mHdr.dwFlags = 0
    ' this next line has caused an error on one user's machine under VB5 - who knows why?
    rc = midiOutPrepareHeader(hMidi, mHdr, LenB(mHdr))
    If rc <> 0 Then
        MsgBox "Prepare rc = " & rc
        Exit Sub
    End If
    ' send long message
    rc = midiOutLongMsg(hMidi, mHdr, LenB(mHdr))
    If rc <> 0 Then
        MsgBox "Send Long Message rc= " & rc
        Exit Sub
    End If
    ' this next line is only required under VB5 IF
    ' you declare the mHdr.lpData as a string
    ' In this new code, mHdr.lpData is a ptr to byte array
    ' and thus this kludge is not required :)
    ' mHdr.dwFlags = 0 ' this line not required anymore
    ' this next line now works under VB5
    rc = midiOutUnprepareHeader(hMidi, mHdr, LenB(mHdr))
    Sleep (Val(Text1.Text))
    If rc <> 0 Then
        MsgBox "Unprepare rc= " & rc
        Exit Sub
    End If
End Sub
Private Sub ReCheck(SyxToRecheck As String)
Const Esa = "0123456789ABCDEF"
'Dim itm As ListItem
Dim Header
Dim ret As Long
List3.Clear
Header = "F0 41 00 14 12"
Dim ins As String
ins = ""
sp = Chr$(32)
'*** Modifica di instruzioni Sysex
'Text2.Text = UCase(Text2.Text)
'Header = Left$(Text2.Text, 14)
For X = 1 To Len(SyxToRecheck) Step 3
    
    CurrentParam = Mid$(SyxToRecheck, X, 2)
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
    'If ("&h" + Param(x).Text) <> (Hex$(Val(Hex2Int(Param(x).Text)))) Then
        
   ' If (Val("&h" + CurrentParam)) > 127 Then
   '     'Par(x).BackColor = &HFF
   '     errore = 1
   '     Exit For
    'End If
Next X
If errore Then GoTo errore

'Chk = Chk + Val("&h" + Text1.Text) + Val("&h" + Text2.Text) + Val("&h" + Text3.Text)

Chk = Chk And 127
Pivot = &H80 - Chk

While (Pivot < 0) Or (Pivot > 127)
Chk = &H80 - (Chk - &H80)
Pivot = Chk
Wend

Checksum = Hex$(Pivot)

If Pivot < 16 Then Checksum = "0" + Checksum
'Set itm = ListView1.ListItems.Add(, , "")
newins = Header + ins + sp + Checksum + sp + "F7"
'itm.SubItems(1) = Mid$(Combo2.List(Combo2.ListIndex), 10, 24)

'Me.Hide
'ListView1.ListItems.Item(ListView1.SelectedItem.Index).Text = newIns
'Text2.Text = newins
'List1.List(List1.ListIndex) = newins
List3.AddItem newins
'Debug.Print " Rechecked: " + newins
errore:
Exit Sub
End Sub

Private Sub List2_Click()
Dim Header

Const Esa = "0123456789ABCDEF"
Dim bytes() As Byte
Dim CharIndex As Integer
Dim ChCur As String
Header = "F0 41 00 14 12"
'RolandChar = " ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890-"
'DumpFile = "C:\prova.dmp"
freef = FreeFile
file_length = FileLen(DumpFile)
ReDim bytes(1 To file_length)
Open DumpFile For Binary Access Read As #freef
Offset = List2.ListIndex * 448
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
'Debug.Print newins
errore:
End Sub



Private Sub Tab1_Click()

If Tab1.SelectedItem.Index = 1 Then
    Frame1.Visible = 1
    Frame2.Visible = 0
    Frame3.Visible = 0
    'Frame4.Visible = 0
End If

If Tab1.SelectedItem.Index = 2 Then
    Frame1.Visible = 0
    Frame2.Visible = 1
    Frame3.Visible = 0
'    Frame4.Visible = 0
End If

If Tab1.SelectedItem.Index = 3 Then
    Frame1.Visible = 0
    Frame2.Visible = 0
    Frame3.Visible = 1
    'Frame4.Visible = 0
End If

'If Tab1.SelectedItem.Index = 4 Then
'    Frame1.Visible = 0
'    Frame2.Visible = 0
'    Frame3.Visible = 0
'    Frame4.Visible = 1
'End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'If KeyAscii = Chr$(13) Then Call Command12_Click
End Sub

