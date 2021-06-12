Attribute VB_Name = "Module1"
Public A, B, C As Single



Sub archivio()
Dim x, xin, yin, xfin, yfin, xfin2, yfin2 As Single

nomefile = "c:\prova_Parabola.dxf"
free = FreeFile
Open nomefile For Output As free
Print #free, 999
Print #free, "DXF creato da Fulcro"


Print #free, 0
Print #free, "SECTION"
Print #free, 2
Print #free, "ENTITIES"

lunghezza = Val(Text8.Text)
Call linea(free, 0, 0, lunghezza, 0, 0) 'asse del campo di disegno

Scalay = Val(Text9.Text)
offset = Val(Text10.Text)
passo = 5
xin = 0
yin = CSng(Parabola(CSng(xin / 100)) / Scalay)
yin2 = CSng(Parabola(CSng((xin / 100) + offset / 100)) / Scalay)
For Ciclo = 0 To (lunghezza - 5) Step 5
xfin = xin + passo
yfin = CSng(Parabola(CSng(xfin / 100)) / Scalay)

Call linea(free, xin, yin, xfin, yfin, 2)
'yfin2 = CSng(Parabola(CSng((xfin / 100) + Offset / 100)) / Scalay)
'Call linea(f, xin, yin2, xfin, yfin2, 3)  ' Diagramma traslato
yin2 = yfin2
xin = xfin
yin = yfin
Next


End Sub

