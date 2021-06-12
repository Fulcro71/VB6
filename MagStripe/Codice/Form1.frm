VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Magstripe Reader - Fulcro2006"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Esci"
      Height          =   495
      Left            =   6240
      TabIndex        =   9
      Top             =   3360
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7200
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salva file"
      Height          =   495
      Left            =   6240
      TabIndex        =   8
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   975
      Left            =   6120
      TabIndex        =   3
      Top             =   840
      Width           =   2175
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   240
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         RTSEnable       =   -1  'True
         SThreshold      =   1
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Traccia:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Numero bytes:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   8175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0006
      Top             =   840
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancella box"
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   2760
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InString As String

Private Sub MSComm1_OnComm()
    Select Case MSComm1.CommEvent
        ' Gestione di eventi o errori presenti nella seguente lista dei case
            

   ' Errori
      Case comEventBreak   ' E' stato ricevuto un Break
      'Print ("EvBreak")
      Case comEventFrame   ' Errore di framing
      'Print ("EvFrame")
      Case comEventOverrun   ' Perdita di dati
      'Print ("EvOverrun")
      Case comEventRxOver   ' Overflow del buffer di ricezione
      'Print ("EvRXOver")
      Case comEventRxParity   ' Errore di parità
      'Print ("EvRXParity")
      Case comEventTxFull   ' Buffer di trasmissione pieno
      'Print ("EvTXFull")
      Case comEventDCB   ' Errore inatteso nella ricezione del DCB
      'Print ("EvDCB")

   ' Eventi
      Case comEvCD   ' Cambio di stato in linea CD
      'Print ("CD")
      Case comEvCTS   ' Cambio di stato in linea CTS
      'Print ("CTS")
      Case comEvDSR   ' Cambio di stato in linea DSR
       'Print ("DSR")
      
      Case comEvRing   ' Cambio nel EventiRing
      'Print ("EvRing")
      Case comEvReceive   ' Ricevuto RThreshold
      
      rx$ = MSComm1.Input
            For i% = 1 To Len(rx$)
                stringa = stringa & Right$("0" + Hex(Asc(Mid$(rx$, i, 1))) & " ", 3)
                raw$ = raw$ + rx$
            Next i%
            raw$ = raw$ + rx$
            Text1.Text = Text1.Text + stringa
            Text2.Text = Text2.Text + rx$
            Text3.Text = Len(Text2.Text)
            TrIndex = Left$(Text2.Text, 1)
            If TrIndex = "%" Then Text4.Text = 1
            If TrIndex = ";" Then Text4.Text = 2
      Case comEvSend   ' Caratteri presenti nel buffer
      Case comEvEOF   ' Un carattere EOF è stato trovato nello stream
    End Select
End Sub

Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Command2_Click()
With CommonDialog1
    .DialogTitle = "Salva file MagStripe"
    .Filter = "*.TXT"
    .FileName = "Carta"
    
    .ShowSave
    
If .FileName = "" Then Exit Sub
Open .FileName For Output As #1: Print #1, Trim(Text2.Text): Close #1
'Open "c:\ascii2byte.txt" For Random As #1: Put #1, , bytBuffer: Close #1

    
End With
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Me.Show
    
    'Usa COM1
    MSComm1.CommPort = 1
    '9600 Baud Rate, Odd Parity, 8 Data e 1 Stop Bit
    MSComm1.Settings = "9600,n,8,1"
    ' istruisce il controllo per la lettura dell'intero buffer quando
    ' viene usato in input
    MSComm1.InputLen = 0
    'Apre la porta
    
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    MSComm1.PortOpen = True
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MSComm1.PortOpen = Not MSComm1.PortOpen
End Sub
