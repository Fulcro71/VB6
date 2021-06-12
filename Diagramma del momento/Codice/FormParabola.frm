VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Crea DXF"
      Height          =   615
      Left            =   5760
      TabIndex        =   35
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Termina"
      Height          =   615
      Left            =   5760
      TabIndex        =   34
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Disegno"
      Height          =   3495
      Left            =   4320
      TabIndex        =   20
      Top             =   120
      Width           =   3375
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   2040
         TabIndex        =   32
         Text            =   "2"
         Top             =   1560
         Width           =   495
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Disegna diagramma traslato"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   2520
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Disegna diagramma principale"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   2160
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   2040
         TabIndex        =   27
         Text            =   "40"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   2160
         TabIndex        =   25
         Text            =   "2"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2040
         TabIndex        =   22
         Text            =   "450"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label18 
         Caption         =   "punti"
         Height          =   255
         Left            =   2640
         TabIndex        =   33
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Intervallo di scansione x:"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   1560
         Width           =   1740
      End
      Begin VB.Label Label16 
         Caption         =   "cm"
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Valore di traslazione:"
         Height          =   195
         Left            =   360
         TabIndex        =   26
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Scala Y del diagramma:  1/"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   1920
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "cm"
         Height          =   195
         Left            =   2640
         TabIndex        =   23
         Top             =   480
         Width           =   210
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Lunghezza del campo:"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   1605
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parabola"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2880
         TabIndex        =   18
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2880
         TabIndex        =   17
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2880
         TabIndex        =   13
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2880
         TabIndex        =   10
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Calcola punti critici"
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000E+00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   6
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   7
         Text            =   "-50"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000E+00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   6
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Text            =   "150"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000E+00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   6
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   5
         Text            =   "-12"
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Intersezioni con l'asse:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1800
         Width           =   1590
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "f(x)="
         Height          =   195
         Left            =   2280
         TabIndex        =   16
         Top             =   3000
         Width           =   300
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "x ="
         Height          =   195
         Left            =   2280
         TabIndex        =   15
         Top             =   2640
         Width           =   210
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Valore di massimo o minimo:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   3000
         Width           =   1965
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "x2="
         Height          =   195
         Left            =   2520
         TabIndex        =   12
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "x1="
         Height          =   195
         Left            =   2520
         TabIndex        =   11
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Punto di massimo o minimo:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "C="
         Height          =   195
         Left            =   2640
         TabIndex        =   4
         Top             =   720
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "B="
         Height          =   195
         Left            =   1560
         TabIndex        =   3
         Top             =   720
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "A="
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Polinomio della parabola: Ax² + Bx + C"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2670
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A, B, C, lunghezza As Single
Dim free As Variant
Private Sub Command1_Click()
Dim prima, x1, seconda, x2 As Double

A = Val(Text1.Text)
B = Val(Text2.Text)
C = Val(Text3.Text)
rad = B ^ 2 - 4 * A * C
If rad < 0 Then
            ret = MsgBox("Attenzione radici complesse.", vbCritical + vbOKOnly, "Soluzioni non reali")
                       
            Exit Sub
            End If
x1 = (-B + Sqr(rad)) / (2 * A)
x2 = (-B - Sqr(rad)) / (2 * A)
Text4.Text = Val(Format$(x1, "##.##"))
Text5.Text = Val(Format$(x2, "##.##"))
PDN = -B / (2 * A)
Text6.Text = Val(Format$(PDN, "##.##"))
Text7.Text = Val(Format$(A * (PDN ^ 2) + B * PDN + C, "##.##"))
End Sub
Function f(num As Single) As Single
f = Val(Format$(num, "##.###"))
End Function

Private Sub Command2_Click()
End Sub


Sub linea(myfile As Variant, xi, yi, xf, yf As Variant, color As Integer)
Dim f
f = myfile

Print #free, "0"
Print #free, "LINE"
Print #free, "8" ' Layer Name
Print #free, "0"
Print #free, "62" 'colore
Print #free, color
Print #free, "10" 'Imposta la x iniziale"
Print #free, xi
Print #free, "20" 'Imposta la y iniziale
Print #free, yi
Print #free, "11" 'x finale
Print #free, xf
Print #free, "21" ' y finale
Print #free, yf
End Sub
Function Parabola(x As Single) As Single
Parabola = -A * (x ^ 2) - B * x - C
End Function

Private Sub Command3_Click()  '**** Trasla ****

Dim x, xin, yin, xfin, yfin, xfin2, yfin2 As Single
lunghezza = Val(Text8.Text)
tratto = Val(Text6.Text) * 100
'Call linea(free, 0, 0, Lunghezza, 0, 0) 'asse del campo di disegno

Scalay = Val(Text9.Text)
FixOffset = Val(Text10.Text)
passo = 1
'xin = 0
xin = 0
yin = CSng(Parabola(CSng(xin / 100)) / Scalay)
'If yin < 0 Then
            'x = x + offset
'            End If
'x = xin + offset

For Ciclo = 0 To (tratto - passo) Step passo
xfin = xin + passo

yfin = CSng(Parabola(CSng(xfin / 100)) / Scalay)

If yfin < 0 Then offset = -Abs(FixOffset) Else offset = Abs(FixOffset)

If xin + offset < 0 Then
                    ' prova
                    Else
                Call linea(free, xin + offset, yin, xfin + offset, yfin, 3)
                End If

xin = xfin
yin = yfin
Next
Call linea(free, xfin, yfin, xfin + offset, yfin, 3)
Call linea(free, xfin, yfin, xfin - offset, yfin, 3)
For Ciclo = tratto To (lunghezza - passo) Step passo
xfin = xin + passo

yfin = CSng(Parabola(CSng(xfin / 100)) / Scalay)
If yfin < 0 Then offset = Abs(FixOffset) Else offset = -Abs(FixOffset)
If xin + offset > lunghezza Then
                    ' prova
                    Else
                Call linea(free, xin + offset, yin, xfin + offset, yfin, 3)
                End If

xin = xfin
yin = yfin
Next
'Call linea(free, xfin, yfin, xfin - offset, yfin, 3)
End Sub


Private Sub Command4_Click()

End
End Sub

Private Sub Command5_Click()
Dim Colore As Integer
Dim x, xin, yin, xfin, yfin, xfin2, yfin2 As Single
nomefile = "c:\Parabola.dxf"
On Error GoTo errore
free = FreeFile
Scalay = Val(Text9.Text)
FixOffset = Val(Text10.Text)
ZeroX1 = Val(Text4.Text)
ZeroX2 = Val(Text5.Text)
lunghezza = Val(Text8.Text)
passo = Val(Text11.Text)

Open nomefile For Output Shared As free
Print #free, 999
Print #free, "DXF creato da Fulcro"


Print #free, 0
Print #free, "SECTION"
Print #free, 2
Print #free, "ENTITIES"
Call linea(free, 0, 0, lunghezza, 0, 0) 'asse del campo di disegno
xin = 0
x = 0
offset = Abs(FixOffset)
'*******************************************************************************
'* Traccia il diagramma originale
'*******************************************************************************


If Check1.Value = 1 Then
Colore = 1
xin = 0

yin = CSng(Parabola(CSng(xin / 100)) / Scalay)
yin2 = CSng(Parabola(CSng((xin / 100) + offset / 100)) / Scalay)
If yin <> 0 Then Call linea(free, 0, 0, xin, yin, Colore)
For Ciclo = 0 To lunghezza - passo Step passo
xfin = xin + passo
yfin = CSng(Parabola(CSng(xfin / 100)) / Scalay)

Call linea(free, xin, yin, xfin, yfin, Colore)

xin = xfin
yin = yfin
Next
If yfin <> 0 Then Call linea(free, xfin, yfin, xfin, 0, Colore)
'*******************************************************************************

End If


 If Check2.Value = 1 Then
 
'*******************************************************************************
'* Traccia il diagramma traslato (fino al punto di derivata nulla)
'*******************************************************************************
tratto = Val(Text6.Text) * 100
If tratto > 0 And tratto < lunghezza Then
                                    ciclo1 = tratto
                                    ciclo2 = lunghezza - tratto
                                    Else
                                    ciclo1 = lunghezza
                                    ciclo2 = 0 'Se il punto di max/min è fuori dall'asse...
                                     End If     ' impedisce di eseguire il 2° ciclo For
Colore = 2
DaChiudere = False
DaChiudereY = 0
minimo = 0
massimo = ciclo1

y = CSng(Parabola(CSng(minimo / 100)) / Scalay)
'***** Chiusura sull'asse della parte iniziale del diagramma:
If y > 0 Then
        Call linea(free, 0, y, offset, y, Colore) 'bordo orizzontale diagramma
        Call linea(free, 0, 0, 0, y, Colore) 'bordo verticale diagramma
        Else
        Call linea(free, 0, 0, 0, CSng(Parabola(CSng(offset / 100)) / Scalay), Colore) 'bordo orizzontale diagramma
        End If

For conto = minimo To massimo Step passo
x = conto

y = CSng(Parabola(CSng(x / 100)) / Scalay)
' In base a y si traccia il diagramma con traslazione positiva o negativa
If y > 0 Then
            'yin = CSng(Parabola(CSng(xin / 100), offset) / Scalay)
            
            xin = x + offset
            If xin >= minimo And xin <= (massimo - passo) Then
            'If xin > (tratto - passo) Then xin = (tratto - passo)
            yin = CSng(Parabola(CSng(x / 100)) / Scalay)
            xfin = x + offset + passo
            yfin = CSng(Parabola(CSng((x + passo) / 100)) / Scalay)
            Call linea(free, xin, yin, xfin, yfin, Colore)
                    End If
          Else
            xin = x - offset
            If xin >= minimo And xin <= (massimo - passo) Then
                    yin = CSng(Parabola(CSng(x / 100)) / Scalay)
                    xfin = x - offset + passo
                    yfin = CSng(Parabola(CSng((x + passo) / 100)) / Scalay)
                    Call linea(free, xin, yin, xfin, yfin, Colore)
                            Else
                                DaChiudere = True
                                y = CSng(Parabola(CSng((x + passo) / 100)) / Scalay)
                                DaChiudereY = y
                                    End If
            End If

DoEvents

Next
If DaChiudere Then Call linea(free, 0, 0, 0, DaChiudereY, Colore)
If massimo > lunghezza Then
                    massimo = lunghezza
                    y = -1
                    End If
Call linea(free, xfin, yfin, massimo, yfin, Colore)
'*******************************************************************************
'* Traccia il diagramma traslato (dal punto di derivata nulla fino a fine campo)
'*******************************************************************************
xin = 0
Colore = 2
minimo = ciclo1
massimo = ciclo1 + ciclo2
If (minimo + offset) <= lunghezza Then Call linea(free, minimo, yfin, minimo + offset, yfin, Colore)
If ciclo2 = 0 Then minimo = massimo + 1 ' Se il punto di max-min è fuori dall'asse non esegue il ciclo for
DaChiudere = False  'Flag di chiusura per diagramma sotto l'asse


For conto = minimo To massimo - passo Step passo
x = conto
y = CSng(Parabola(CSng(x / 100)) / Scalay)
' In base a y si traccia il diagramma con traslazione positiva o negativa
If y < 0 Then
            xin = x + offset
            If xin >= minimo And xin < massimo Then
            yin = CSng(Parabola(CSng(x / 100)) / Scalay)
            xfin = x + offset + passo
            yfin = CSng(Parabola(CSng((x + passo) / 100)) / Scalay)
            
            Call linea(free, xin, yin, xfin, yfin, 2)
                Else
                       DaChiudere = True
                                    End If
          Else
          xin = x - offset
            If xin > minimo And xin < massimo Then
                 DaChiudere = True
                 yin = CSng(Parabola(CSng(x / 100)) / Scalay)
                 DaChiudereY = yin
                 xfin = x - offset + passo
                 yfin = CSng(Parabola(CSng((x + passo) / 100)) / Scalay)
                 Call linea(free, xin, yin, xfin, yfin, 2)
                            
                            End If
                    
            End If
DoEvents
Next

'*******************************************************************************
'* Chiusura dei diagrammi sull'asse
'*******************************************************************************
y = CSng(Parabola(CSng(massimo / 100)) / Scalay) 'y finale del diagramma
If y > 0 Then   'Se il diagramma è sulla parte superiore dell'asse:
            '*******Linea Verticale dalla y del diagramma fino all'asse
            Call linea(free, massimo, 0, massimo, yfin, Colore)
            '*******Linea orizzontale dalla fine del diagramma fino a fine campo:
            'Call linea(free, massimo, yfin, massimo - offset + passo, yfin, Colore)
            Call linea(free, massimo, yfin, massimo - offset, yfin, Colore)
            Else    'Se il diagramma sta sotto l'asse:
            '*******Linea Verticale dalla fine del diagramma fino all'asse
            Call linea(free, lunghezza, 0, lunghezza, yfin, Colore)
            End If
            
'**** Chiusura dei diagrammi sotto l'asse versol'asse:
If DaChiudere Then Call linea(free, lunghezza, 0, lunghezza, CSng(Parabola(CSng((lunghezza - offset) / 100)) / Scalay), Colore)
End If
'********** Scrittura della fine del file DXF e chiusura del file********
Print #free, 0
Print #free, "ENDSEC"

'Fine File
Print #free, 0
Print #free, "EOF"

Close #free
Exit Sub
errore:
ret = MsgBox("Errore nella creazione del file: il file potrebbe essere già in uso.", vbCritical + vbOKOnly, "Errore!")
End Sub

Private Sub trasla()
' Era la routine associata all'evento Click del pulsante Command3
Dim x, xin, yin, xfin, yfin, xfin2, yfin2 As Single
lunghezza = Val(Text8.Text)
tratto = Val(Text6.Text) * 100
'Call linea(free, 0, 0, Lunghezza, 0, 0) 'asse del campo di disegno

Scalay = Val(Text9.Text)
FixOffset = Val(Text10.Text)
passo = 1
'xin = 0
xin = 0
yin = CSng(Parabola(CSng(xin / 100)) / Scalay)
'If yin < 0 Then
            'x = x + offset
'            End If
'x = xin + offset

For Ciclo = 0 To (tratto - passo) Step passo
xfin = xin + passo

yfin = CSng(Parabola(CSng(xfin / 100)) / Scalay)

If yfin < 0 Then offset = -Abs(FixOffset) Else offset = Abs(FixOffset)

If xin + offset < 0 Then
                    ' prova
                    Else
                Call linea(free, xin + offset, yin, xfin + offset, yfin, 3)
                End If

xin = xfin
yin = yfin
Next
Call linea(free, xfin, yfin, xfin + offset, yfin, 3)
Call linea(free, xfin, yfin, xfin - offset, yfin, 3)
For Ciclo = tratto To (lunghezza - passo) Step passo
xfin = xin + passo

yfin = CSng(Parabola(CSng(xfin / 100)) / Scalay)
If yfin < 0 Then offset = Abs(FixOffset) Else offset = -Abs(FixOffset)
If xin + offset > lunghezza Then
                    ' prova
                    Else
                Call linea(free, xin + offset, yin, xfin + offset, yfin, 3)
                End If

xin = xfin
yin = yfin
Next
'Call linea(free, xfin, yfin, xfin - offset, yfin, 3)

End Sub
Private Sub Parabola_()
' Era la routine associata all'evento Click del pulsante Command2

Dim PrevOffset
Dim x, xin, yin, xfin, yfin, xfin2, yfin2 As Single
nomefile = "c:\prova_Parabola.dxf"
On Error GoTo errore
free = FreeFile
Open nomefile For Output Shared As free
Print #free, 999
Print #free, "DXF creato da Fulcro"


Print #free, 0
Print #free, "SECTION"
Print #free, 2
Print #free, "ENTITIES"
lunghezza = Val(Text8.Text)
passo = 2
Call linea(free, 0, 0, lunghezza, 0, 0) 'asse del campo di disegno
'******************************************************************
'* Tracciamento del diagramma principale
'******************************************************************
If Check1.Value = 1 Then
Scalay = Val(Text9.Text)
offset = Val(Text10.Text)

xin = 0
yin = CSng(Parabola(CSng(xin / 100)) / Scalay)
yin2 = CSng(Parabola(CSng((xin / 100) + offset / 100)) / Scalay)
If yin <> 0 Then Call linea(free, 0, 0, xin, yin, 2)
For Ciclo = 0 To (lunghezza - passo) Step passo
xfin = xin + passo
yfin = CSng(Parabola(CSng(xfin / 100)) / Scalay)

Call linea(free, xin, yin, xfin, yfin, 2)
'yfin2 = CSng(Parabola(CSng((xfin / 100) + Offset / 100)) / Scalay)
'Call linea(f, xin, yin2, xfin, yfin2, 3)  ' Diagramma traslato
yin2 = yfin2
xin = xfin
yin = yfin
Next
If yfin <> 0 Then Call linea(free, xfin, yfin, xfin, 0, 2)
End If


If Check2.Value = 1 Then
'******************************************************************
'* Disegna tratto traslato fino all'eventuale punto di derivata nulla
'******************************************************************
tratto = Val(Text6.Text) * 100
Scalay = Val(Text9.Text)
FixOffset = Val(Text10.Text)

xin = 0
yin = CSng(Parabola(CSng(xin / 100)) / Scalay)
'************************** Bordo sinistro del diagramma traslato ************
If yin < 0 Then
            Call linea(free, 0, 0, 0, CSng(Parabola(CSng((xin + FixOffset) / 100)) / Scalay), 3)
            Else
            Call linea(free, 0, 0, 0, yin, 3)
            Call linea(free, 0, yin, FixOffset, yin, 3)
            'rem
            End If
'*****************************************************************************
If tratto > lunghezza Or tratto < 0 Then tratto = lunghezza
PrevOffset = FixOffset

For Ciclo = 0 To (tratto - passo) Step passo

xfin = xin + passo

yfin = CSng(Parabola(CSng(xfin / 100)) / Scalay)

If yin < 0 Then offset = -Abs(FixOffset) Else offset = Abs(FixOffset)

If xin + offset > 0 Then
                    If (offset <> PrevOffset) Then
                    Call linea(free, prevxin + offset, prevyin, xin + offset, yin, 4)
                    End If
                    Call linea(free, xin + offset, yin, xfin + offset, yfin, 3)
                    
                    Else
                    'If yin <> 0 Then Call linea(free, 0, 0, 0, yin, 3)
                End If
PrevOffset = offset
prevxin = xin
prevyin = yin
xin = xfin
yin = yfin
Next
Call linea(free, xfin, yfin, xfin + offset, yfin, 3)
'******************************************************************
'* Disegna tratto traslato dall'eventuale punto di derivata nulla
'******************************************************************
If tratto < lunghezza Then
            
            Call linea(free, xfin, yfin, xfin - offset, yfin, 3)
            End If
PrevOffset = FixOffset
            
For Ciclo = tratto To (lunghezza - passo) Step passo
xfin = xin + passo

yfin = CSng(Parabola(CSng(xfin / 100)) / Scalay)
If yfin < (passo / 2) Then
                offset = Abs(FixOffset)
                Else
                offset = -Abs(FixOffset)
                End If
                
If xin + offset < lunghezza Then
                    If (offset <> PrevOffset) Then
                    Call linea(free, prevxin + offset, prevyin, prevxin + offset + passo, CSng(Parabola(CSng((prevxin + passo) / 100)) / Scalay), 4)
                    End If
                    Call linea(free, xin + offset, yin, xfin + offset, yfin, 3)
                    Else
                    'nop
                End If
PrevOffset = offset
prevxin = xin
prevyin = yin
xin = xfin
yin = yfin
Next
If yfin < 0 Then
            Call linea(free, lunghezza, 0, lunghezza, CSng(Parabola(CSng((xfin) / 100)) / Scalay), 3)
            Else
            Call linea(free, lunghezza, 0, lunghezza, yfin, 3)
            Call linea(free, lunghezza, yfin, lunghezza - FixOffset, yfin, 3)
            'rem
            End If
End If
Print #free, 0
Print #free, "ENDSEC"

'Fine File
Print #free, 0
Print #free, "EOF"

Close #free
Exit Sub
errore:
ret = MsgBox("Errore nella creazione del file: il file potrebbe essere già in uso.", vbCritical + vbOKOnly, "Errore!")

End Sub

