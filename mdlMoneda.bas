Attribute VB_Name = "mdlMoneda"
Option Compare Database
Option Explicit

Dim cTexto As String

Public Function NumLetras(ByVal Numero As Double, ByVal Mayusculas As Integer, Moneda As Integer) As String
  Dim NumTmp As String
  Dim c01 As Integer
  Dim c02 As Integer
  Dim POS As Integer
  Dim dig As Integer
  Dim cen As Integer
  Dim dec As Integer
  Dim uni As Integer
  Dim letra1 As String
  Dim letra2 As String
  Dim letra3 As String
  Dim Leyenda As String
  Dim Leyenda1 As String
  Dim TFNumero As String

  If Numero < 0 Then Numero = Abs(Numero)

  NumTmp = Format(Numero, "000000000000000.00")        'Le da un formato fijo
  c01 = 1
  POS = 1
  TFNumero = ""
  Do While c01 <= 5
    c02 = 1
    Do While c02 <= 3
      dig = Val(Mid(NumTmp, POS, 1))
     Select Case c02
        Case 1: cen = dig
        Case 2: dec = dig
        Case 3: uni = dig
      End Select
      c02 = c02 + 1
      POS = POS + 1
    Loop
    letra3 = Centena(uni, dec, cen)
    letra2 = Decena(uni, dec)
    letra1 = Unidad(uni, dec)

    Select Case c01
      Case 1
        If cen + dec + uni = 1 Then
          Leyenda = "Billon "
        ElseIf cen + dec + uni > 1 Then
          Leyenda = "Billones "
        End If
      Case 2
        If cen + dec + uni >= 1 And Val(Mid(NumTmp, 7, 3)) = 0 Then
          Leyenda = "Mil Millones "
        ElseIf cen + dec + uni >= 1 Then
          Leyenda = "Mil "
        End If
      Case 3
        If cen + dec = 0 And uni = 1 Then
          Leyenda = "Millon "
        ElseIf cen > 0 Or dec > 0 Or uni > 1 Then
          Leyenda = "Millones "
        End If
      Case 4
        If cen + dec + uni >= 1 Then
          Leyenda = "Mil "
        End If
      Case 5
        If cen + dec + uni >= 1 Then
          Leyenda = ""
        End If
      End Select

      c01 = c01 + 1

      TFNumero = TFNumero + letra3 + letra2 + letra1 + Leyenda

      Leyenda = ""
      letra1 = ""
      letra2 = ""
      letra3 = ""

  Loop

  If Val(NumTmp) = 0 Or Val(NumTmp) < 1 Then
    Leyenda1 = "Cero Pesos "
  ElseIf Val(NumTmp) = 1 Or Val(NumTmp) < 2 Then
    If Moneda = 1 Then
        Leyenda1 = "Peso "
        Else
        Leyenda1 = "Dólares "
    End If
    
  ElseIf Val(Mid(NumTmp, 4, 12)) = 0 Or Val(Mid(NumTmp, 10, 6)) = 0 Then
    If Moneda = 1 Then
        Leyenda1 = "de Pesos "
        Else
        Leyenda1 = "de Dólares "
    End If
  Else
  If Moneda = 1 Then
    Leyenda1 = "Pesos "
    Else
    Leyenda1 = "Dólares "
End If
  End If
    If Moneda = 1 Then
        TFNumero = "(" & TFNumero & Leyenda1 & Mid(NumTmp, 17) & "/100 MCTE.)"
        Else
        TFNumero = "( " & TFNumero & Leyenda1 & Mid(NumTmp, 17) & "/100 CENTS.)"
    End If
  If Mayusculas = 1 Then
    TFNumero = UCase(TFNumero)
  Else
    TFNumero = LCase(TFNumero)
  End If
  NumLetras = TFNumero
End Function

Private Function Centena(ByVal uni As Integer, ByVal dec As Integer, _
                         ByVal cen As Integer) As String
  Select Case cen
    Case 1
      If dec + uni = 0 Then
        cTexto = "cien "
      Else
        cTexto = "ciento "
      End If
    Case 2: cTexto = "doscientos "
    Case 3: cTexto = "trescientos "
    Case 4: cTexto = "cuatrocientos "
    Case 5: cTexto = "quinientos "
    Case 6: cTexto = "seiscientos "
    Case 7: cTexto = "setecientos "
    Case 8: cTexto = "ochocientos "
    Case 9: cTexto = "novecientos "
    Case Else: cTexto = ""
  End Select

  Centena = cTexto
  cTexto = ""
End Function

Private Function Decena(ByVal uni As Integer, ByVal dec As Integer) As String
  Select Case dec
    Case 1
      Select Case uni
        Case 0: cTexto = "diez "
        Case 1: cTexto = "once "
        Case 2: cTexto = "doce "
        Case 3: cTexto = "trece "
        Case 4: cTexto = "catorce "
        Case 5: cTexto = "quince "
        Case 6 To 9: cTexto = "dieci"
      End Select
    Case 2
      If uni = 0 Then
        cTexto = "veinte "
      ElseIf uni > 0 Then
        cTexto = "veinti"
      End If
    Case 3: cTexto = "treinta "
    Case 4: cTexto = "cuarenta "
    Case 5: cTexto = "cincuenta "
    Case 6: cTexto = "sesenta "
    Case 7: cTexto = "setenta "
    Case 8: cTexto = "ochenta "
    Case 9: cTexto = "noventa "
    Case Else: cTexto = ""
  End Select

  If uni > 0 And dec > 2 Then cTexto = cTexto + "y "

  Decena = cTexto
  cTexto = ""
End Function

Private Function Unidad(ByVal uni As Integer, ByVal dec As Integer) As String
  If dec <> 1 Then
    Select Case uni
      Case 1: cTexto = "un "
      Case 2: cTexto = "dos "
      Case 3: cTexto = "tres "
      Case 4: cTexto = "cuatro "
      Case 5: cTexto = "cinco "
    End Select
  End If

  Select Case uni
    Case 6: cTexto = "seis "
    Case 7: cTexto = "siete "
    Case 8: cTexto = "ocho "
    Case 9: cTexto = "nueve "
  End Select

  Unidad = cTexto
  cTexto = ""
End Function

                            


