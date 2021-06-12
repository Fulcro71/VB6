Attribute VB_Name = "Module1"
Public Afc, Aft, CFerroC, CFerroT As Single
Public Rck, fck, fyk, ftk, fcd, fyd, ftd As Single
Public eyd, ec, es, es1, xi, beta, k, delta, b, d, yn, RhoS, RhoS1 As Single
Public SigmaC, SigmaS, SigmaS1 As Single
Const ecu = 0.0035
Const esu = 0.01

'********************************************************
'* Da rilevare dal Registro o da file di configurazione
'********************************************************
Const Alfa = 0.85
Const GammaC = 1.6 ' Coeff. di sicurezza per il calcestruzzo
Const GammaS = 1.15 ' Coeff. di sicurezza per l'acciaio
'Const CferroT = 4   ' Copriferro teso [cm]         Variato il 29-11-2005
'Const CFerroC = 4   ' Copriferro compresso  [cm]
Public Const Eacciaio = 200000   ' Mod. elastico ACCIAIO [N/mm^2]


Sub Assign()
CFerroT = CSng(Form1.Text7.Text)    'Aggiunti il 29-11-2005
CFerroC = CFerroT                   ' Prima era costante, ugule a 4 cm.

fcd = Format$(fck / GammaC, "##.##")
fyd = Format$(fyk / GammaS, "##.##")
ftd = Format$(ftk / GammaS, "##.##")
'eyd = Val(Format$(fyd / Eacciaio, "#.#####"))
eyd = fyd / Eacciaio
d = CSng(Form1.Text6.Text)
delta = CFerroC / d
'Es = CSng(form1.text12.Text)
b = CSng(Form1.Text5.Text)
RhoS = Format$(Aft / (b * d), "#.#####")
RhoS1 = Format$(Afc / (b * d), "#.#####")
'Form1.Text7.Text = 3
Form1.Text13.Text = fcd
Form1.Text14.Text = fyd
Form1.Text15.Text = ftd
Form1.Text16.Text = Format$(eyd, "##.#####")
Form1.Text29.Text = Format$((Alfa * fcd), "##.##")
Form1.Text17.Text = Format(delta, "#.#####")
Form1.Text18.Text = Format(RhoS, "#.#####")
Form1.Text19.Text = Format(RhoS1, "#.#####")
Form1.Command2.Enabled = True

Debug.Print "fcd= " & fcd
Debug.Print "fyd= " & fyd
Debug.Print "ftd= " & ftd
Debug.Print "eyd= " & eyd
Debug.Print "delta= " & delta
Debug.Print Aft
Debug.Print Afc

End Sub

Function Campo2() As Single
Dim xi_min, xi_max, t As Single
xi_min = 0
xi_max = 0.259
xi = xi_min
es = esu
done = False
Campo2 = -1
winCap = Form1.Caption
Form1.Caption = Form1.Caption + " - Campo 2 ...."
Threshold = 0.001
SigmaC = Alfa * fcd     'Il calcestruzzo non è in crisi ma si considera comunque tale...
While (Not done) And (xi < xi_max)

ec = (esu * xi) / (1 - xi)

If xi < delta Then
            es1 = esu * (delta - xi) / (1 - xi)
            Else
            es1 = esu * (xi - delta) / (1 - xi)
        End If

 'If es < eyd Then
 '               SigmaS = Eacciaio * es
 '           Else
 '               'Normativa Italiana
 '               ' SigmaS=fyd
 '               'Normativa Europea EC2
 '                SigmaS = ((ftd - fyd) * (es - eyd) / (0.01 + eyd)) + fyd
 es = esu
    SigmaS = ((ftd - fyd) * (es - eyd) / (0.01 - eyd)) + fyd
    '    End If
If es1 < eyd Then
                SigmaS1 = Eacciaio * es1
            Else
                'Normativa Italiana
                ' SigmaS1=fyd
                'Normativa Europea EC2
                 SigmaS = ((ftd - fyd) * (es1 - eyd) / (0.01 - eyd)) + fyd
        End If

'Debug.Print "beta=" & beta & "    " & "k=" & k
'Debug.Print
DoEvents
eupc = ec / ecu
    beta = (1.6 - 0.8 * eupc) * eupc
    k = 0.33 + 0.07 * eupc
'*************** Normativa Italiana ******************
'test = -SigmaC * beta * xi - SigmaS1 * RhoS1 + fyd * RhoS

'*************** Normativa Europea EC2 ***************
test = -SigmaC * beta * xi - SigmaS1 * RhoS1 + ftd * RhoS

t = test
If Abs(t) < Threshold Then
    Threshold = t
    XiThreshold = xi
    Campo2 = XiThreshold
    

Form1.Text20.Text = Format(ec, "##.#####")
Form1.Text21.Text = Format(es, "##.#####")
Form1.Text22.Text = Format(es1, "##.#####")
Form1.Text23.Text = Format(SigmaC, "##.##")
Form1.Text24.Text = Format(SigmaS, "##.##")
Form1.Text25.Text = Format(SigmaS1, "##.##")
    End If
         xi = xi + 0.00001
Wend
Campo2 = XiThreshold
Form1.Caption = winCap

End Function

Function Campo3() As Single             '  ***
Dim xi_min, xi_max As Single            ' *   *
xi_min = 0.259                          '    *
xi_max = 0.0035 / (0.0035 + eyd)        ' *   *
xi = xi_min                             '  ***
ec = ecu
es1 = 0
SigmaS1 = 0
done = False
Campo3 = -1
Threshold = 0.001
beta = 0.8
k = 0.4
SigmaC = Alfa * fcd
winCap = Form1.Caption
Form1.Caption = Form1.Caption + " - Campo 3 ...."
While (Not done) And (xi < xi_max)
es = (ecu * (1 - xi)) / xi
es1 = (ecu * (xi - delta)) / xi

 'If es < eyd Then
 '               SigmaS = Eacciaio * es
 '           Else
 '                '*********Normativa Italiana*********
 '                'SigmaS = fyd
 '
 '                '*********Normativa Europea EC2*********
                 SigmaS = ((ftd - fyd) * (es - eyd) / (0.01 - eyd)) + fyd
'        End If
If es1 < eyd Then
                 SigmaS1 = Eacciaio * es1
            Else
                 '*********Normativa Italiana*********
                 'SigmaS1 = fyd
                 
                 '*********Normativa Europea EC2*********
                 SigmaS1 = ((ftd - fyd) * (es1 - eyd) / (0.01 - eyd)) + fyd
        End If


DoEvents
test = -Alfa * fcd * beta * xi - SigmaS1 * RhoS1 + SigmaS * RhoS

't = Format$(test, "##.####")
t = test
If Abs(t) < Threshold Then
    Threshold = t
    XiThreshold = xi
    Campo3 = XiThreshold
    Form1.Text20.Text = Format(ec, "##.#####")
    Form1.Text21.Text = Format(es, "##.#####")
    Form1.Text22.Text = Format(es1, "##.#####")
    Form1.Text23.Text = Format(SigmaC, "##.##")
    Form1.Text24.Text = Format(SigmaS, "##.##")
    Form1.Text25.Text = Format(SigmaS1, "##.##")
    End If
         xi = xi + 0.00001
Wend
'Campo3 = XiThreshold
Form1.Caption = winCap
End Function

Function Campo4() As Single             '    **
Dim xi_min, xi_max As Single            '   * *
xi_min = 0.0035 / (0.0035 + eyd)        '  *  *
xi_max = 1                              ' ******
'xi_max = 0.3                           '     *
xi = xi_min
ec = ecu
done = False
Campo4 = -1
winCap = Form1.Caption
Threshold = 0.001
SigmaC = Alfa * fcd
beta = 0.8
k = 0.4
Form1.Caption = Form1.Caption + " - Campo 4 ...."
While (Not done) And (xi < xi_max)
es = ecu * (1 - xi) / xi
es1 = ecu * (xi - delta) / xi

If es < eyd Then
                SigmaS = Eacciaio * es
                Else
            '*********Normativa Italiana*********
                'SigmaS1 = fyd
                 '*********Normativa Europea EC2*********
                 SigmaS = (((ftd - fyd) * (es - eyd)) / (0.01 - eyd)) + fyd
       End If
If es1 < eyd Then
                SigmaS1 = Eacciaio * es1
                Else
            '*********Normativa Italiana*********
                'SigmaS1 = fyd
                 '*********Normativa Europea EC2*********
                 SigmaS1 = (((ftd - fyd) * (es1 - eyd)) / (0.01 - eyd)) + fyd
       End If

DoEvents
test = CSng(-Alfa * fcd * beta * xi - SigmaS1 * RhoS1 + SigmaS * RhoS)

t = test
If Abs(t) < Threshold Then
       Threshold = t
    XiThreshold = xi
    Campo4 = XiThreshold
Form1.Text20.Text = Format(ec, "##.#####")
Form1.Text21.Text = Format(es, "##.#####")
Form1.Text22.Text = Format(es1, "##.#####")
Form1.Text23.Text = Format(SigmaC, "##.##")
Form1.Text24.Text = Format(SigmaS, "##.##")
Form1.Text25.Text = Format(SigmaS1, "##.##")
    End If
    xi = xi + 0.00001
Wend
Form1.Caption = winCap
End Function
Function Formatta(num As Double) As Single
n = num

n = Val(Format$(n, "##.########"))

Formatta = n
End Function
Sub DefStatus(X As Single)
yn = Val(Format$((X * d), "##.##"))
Form1.Text27.Text = (Format$(X, "#.#####"))
Form1.Text28.Text = yn
End Sub
Sub Mrd()
mu = (SigmaC * beta * (b * 10) * (yn * 10) * (d * 10 - k * yn * 10) + Form1.Text25.Text * (Afc * 100) * (d * 10 - CFerroT * 10)) / (10 ^ 6)
Form1.Text26.Text = Format$(mu, "###.###")

End Sub

Sub Resetta()
With Form1

.Command2.Enabled = 0
'Form1.Text12.Text = ""
.Text13.Text = ""
.Text14.Text = ""
.Text15.Text = ""
.Text16.Text = ""
.Text17.Text = ""
.Text18.Text = ""
.Text19.Text = ""
.Text20.Text = ""
.Text21.Text = ""
.Text22.Text = ""
.Text23.Text = ""
.Text24.Text = ""
.Text25.Text = ""
.Text26.Text = ""
.Text27.Text = ""
.Text28.Text = ""
.Text29.Text = ""
End With
End Sub
