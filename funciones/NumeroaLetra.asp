<%
Dim xcen(9) 'centenas
Dim xdec(9) 'decenas
Dim xuni(9) 'unidades
Dim xexc(6) 'except
Dim ceros(9)

Function CONVERTIR(pnumero)

Dim letras
Dim i
Dim c
Dim j
Dim xnumero
Dim xnum
Dim num
Dim digito
Dim numero_ent
Dim entero
Dim decimales
Dim temp
  
  xcen(2) = "dosc"
  xcen(3) = "tresc"
  xcen(4) = "cuatrosc"
  xcen(5) = "quin"
  xcen(6) = "seisc"
  xcen(7) = "setec"
  xcen(8) = "ochoc"
  xcen(9) = "novec"
  xdec(2) = "veinti"
  xdec(3) = "trei"
  xdec(4) = "cuare"
  xdec(5) = "cincue"
  xdec(6) = "sese"
  xdec(7) = "sete"
  xdec(8) = "oche"
  xdec(9) = "nove"
  xuni(1) = "uno"
  xuni(2) = "dos"
  xuni(3) = "tres"
  xuni(4) = "cuatro"
  xuni(5) = "cinco"
  xuni(6) = "seis"
  xuni(7) = "siete"
  xuni(8) = "ocho"
  xuni(9) = "nueve"
  xexc(1) = "diez"
  xexc(2) = "once"
  xexc(3) = "doce"
  xexc(4) = "trece"
  xexc(5) = "catorce"
  xexc(6) = "quince"
  ceros(1) = "0"
  ceros(2) = "00"
  ceros(3) = "000"
  ceros(4) = "0000"
  ceros(5) = "00000"
  ceros(6) = "000000"
  ceros(7) = "0000000"
  ceros(8) = "00000000"
  
  c = 1
  i = 1
  j = 0
  
  xnumero = cStr(pnumero)
If Cdbl(LTrim(RTrim(pnumero))) < 999999999.99 Then
    numero_ent = Cdbl(Int(pnumero))
    If Len(numero_ent) < 9 Then
        numero_ent = ceros(9 - Len(numero_ent)) & numero_ent
    End If
    entero = Cdbl(Int(numero_ent))
    decimales = (Cdbl(xnumero) - entero) * 100
    
	Do While i < 8
        temp = 0
        num = Cdbl(Mid(numero_ent, i, 3))
        xnum = Mid(numero_ent, i, 3)
        digito = Cdbl(Mid(xnum, 1, 1))
        
        '/* analizo el numero entero de a 3 */
        If xnum = "000" Then
            j = 0
        Else
        	j = 1
            If digito > 1 Then
                letras = letras & xcen(digito) & "ientos "
            End If
            If Mid(xnum, 1, 1) = "1" And Mid(xnum, 2, 2) <> "00" Then
                letras = letras & "ciento "
            ElseIf Mid(xnum, 1, 1) = "1" Then
                letras = letras & "cien "
            End If
  
  			'/* analisis de las decenas */
            digito = Cdbl(Mid(xnum, 2, 1))
            If digito > 2 And Mid(xnum, 3, 1) = "0" Then
                letras = letras & xdec(digito) & "nta "
				temp = 1
            End If
            
			If digito > 2 And Mid(xnum, 3, 1) <> "0" Then
                letras = letras & xdec(digito) & "nta y "
				
            End If
            
			If digito = 2 And Mid(xnum, 3, 1) = "0" Then
                letras = letras & "veinte "
				temp = 1
            ElseIf digito = 2 And Mid(xnum, 3, 1) <> "0" Then
                letras = letras & "veinti"
				
            End If
            
			If digito = 1 And Mid(xnum, 3, 1) >= "6" Then
                letras = letras & "dieci"
            ElseIf digito = 1 And Mid(xnum, 3, 1) < "6" Then
                letras = letras & xexc(Cdbl(Mid(xnum, 3, 1) + 1))
            	temp = 1
			End If
        End If
   
   		if temp = 0 then
   	'/* analisis del ultimo digito */
        digito = Cdbl(Mid(xnum, 3, 1))
            	If ((c = 1) Or (c = 2)) And xnum = "001" Then
                	letras = letras & "un"
            	Else
                	If ((c = 1) Or (c = 2)) And xnum >= "020" And Mid(xnum, 3, 1) = "1" Then
                    	letras = letras & "un"
                	Else
                    	If digito <> 0 Then
                        	letras = letras & xuni(digito)
                    	End If
                	End If
            	End If
		end if
  
  If j = 1 And i = 1 And xnum = "001" And c = 1 Then
    letras = letras & " millon "
  ElseIf j = 1 And i = 1 And xnum <> "001" And c = 1 Then
    letras = letras & " millones "
  ElseIf j = 1 And i = 4 And c = 2 Then
    letras = letras & " mil "
  End If
  i = i + 3
  c = c + 1
  Loop
  If letras = "" Then
  letras = "cero "
  End If
  If decimales <> 0 Then
    decimales = Round(decimales)
    
    letras = letras & " con " & CStr(decimales) & "/100"
  End If
  
End If

'EN ESTA VARIABLE SESSION SE GUARDA EL NUMERO EN LETRAS 
Session("Valor") = letras
CONVERTIR=letras
End Function

%>