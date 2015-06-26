Attribute VB_Name = "mBarCode128"
Option Explicit

Private Const MAX_PATH As Long = 260

Private Declare Function GetTempPath Lib "kernel32" _
        Alias "GetTempPathA" (ByVal nBufferLenght _
        As Long, ByVal lpBuffer As String) As Long

Public Function calculaDVCodBar128C(ByVal sequencia As String, ByRef pDV As Integer) As Boolean
    
    On Error GoTo trataErro
        
    Dim intMultiplicador As Integer, intResto As Integer, intresultado As Integer
    Dim varfinalLoop As Integer, varPosicao As Integer
    Dim varNumero As Double, varTotalTemp As Double
                        
    intMultiplicador = 2
    varPosicao = 1
    
    varfinalLoop = IIf((Len(sequencia) Mod 2 = 0), (Len(sequencia) / 2), (Len(sequencia) / 2) + 1)
    
    varTotalTemp = 105  'Valor do START
    For intMultiplicador = 1 To varfinalLoop
                                               
        varNumero = Mid(sequencia, varPosicao, 2)
        varPosicao = varPosicao + 2
        
        varTotalTemp = varTotalTemp + (varNumero * intMultiplicador)
                
    Next
        
    intResto = varTotalTemp Mod 103
             
    If intResto > 104 Then
        'O conjunto de Caracteres do Código de Barras CODE-128C só vai até 104 posições
        calculaDVCodBar128C = False
    End If
    
    pDV = intResto
    calculaDVCodBar128C = True
    Exit Function
                

trataErro:
    
    
    calculaDVCodBar128C = False

End Function

Public Function DesenhaCodigoBarras128C(ByVal pSequenciaCodBar As String, ByVal pObjeto As Control, ByVal pExibirNumeroCodBarras As Boolean, ByVal pTamanhoFonte As Integer, ByVal pHeightCodBarras As Integer, ByRef pMensagemErro As String) As Boolean
        
    'On Error GoTo trataErro
    
    '=============== DANFE (NFe)===========================
    
    'Margem clara  ===  Start C  ===  Dados representados  ===  DV  ===  Stop  ===  Margem clara
    
    '>>>>> Margem Clara: espaço claro que não contém nenhuma marca legível por máquina,
    'localizado à esquerda e à direita do código, a fim de evitar interferência na decodificação
    'da simbologia. A margem clara é chamada também de "área livre", "zona de silêncio" ou "margem de silêncio".
    
    '>>>>> Start C: inicia a codificação dos dados CODE-128C de acordo com o conjunto de
    'caracteres. O Start C não representa nenhum caractere.
    
    '>>>>> Dados representados: caracteres representados no código de barras.
    
    '>>>>> DV: dígito verificador da simbologia.
    
    '>>>>> Stop: caractere de parada que indica o final do código ao leitor óptico.
             
    'Arquivo referência: Manual_Integracao_Contribuinte_versao_4.01-NT2009.006
    '======================================================
            
    Dim varPadraoCodBarras128C() As String
    Dim bc(106) As String
    Dim varDV As Integer
    Dim varSeqCodigoBarras As String
    Dim dw As Integer, th As Integer, tw As Integer, xpos As Integer, n As Integer
    Dim y1 As Integer, y2 As Integer, i As Integer, varPosicao As Integer
    Dim new_string As String, c As String
    Dim varDesenharBarra As Boolean
           
    '=======================
    'Combinação de barras: B = barra preta e S = espaço (barra branca)
    
    'Valor
    'B S B S B S B S B S B S B
    '2 1 1 2 3 2 2 3 3 1 1 1 2
    
    'CODE C = B S B S B S
    bc(0) = "212222"
    bc(1) = "222122"
    bc(2) = "222221"
    bc(3) = "121223"
    bc(4) = "121322"
    bc(5) = "131222"
    bc(6) = "122213"
    bc(7) = "122312"
    bc(8) = "132212"
    bc(9) = "221213"
    bc(10) = "221312"
    bc(11) = "231212"
    bc(12) = "112232"
    bc(13) = "122132"
    bc(14) = "122231"
    bc(15) = "113222"
    bc(16) = "123122"
    bc(17) = "123221"
    bc(18) = "223211"
    bc(19) = "221132"
    bc(20) = "221231"
    bc(21) = "213212"
    bc(22) = "223112"
    bc(23) = "312131"
    bc(24) = "311222"
    bc(25) = "321122"
    bc(26) = "321221"
    bc(27) = "312212"
    bc(28) = "322112"
    bc(29) = "322211"
    bc(30) = "212123"
    bc(31) = "212321"
    bc(32) = "232121"
    bc(33) = "111323"
    bc(34) = "131123"
    bc(35) = "131321"
    bc(36) = "112313"
    bc(37) = "132113"
    bc(38) = "132311"
    bc(39) = "211313"
    bc(40) = "231113"
    bc(41) = "231311"
    bc(42) = "112133"
    bc(43) = "112331"
    bc(44) = "132131"
    bc(45) = "113123"
    bc(46) = "113321"
    bc(47) = "133121"
    bc(48) = "313121"
    bc(49) = "211331"
    bc(50) = "231131"
    bc(51) = "213113"
    bc(52) = "213311"
    bc(53) = "213131"
    bc(54) = "311123"
    bc(55) = "311321"
    bc(56) = "331121"
    bc(57) = "312113"
    bc(58) = "312311"
    bc(59) = "332111"
    bc(60) = "314111"
    bc(61) = "221411"
    bc(62) = "431111"
    bc(63) = "111224"
    bc(64) = "111422"
    bc(65) = "121124"
    bc(66) = "121421"
    bc(67) = "141122"
    bc(68) = "141221"
    bc(69) = "112214"
    bc(70) = "112412"
    bc(71) = "122114"
    bc(72) = "122411"
    bc(73) = "142112"
    bc(74) = "142211"
    bc(75) = "241211"
    bc(76) = "221114"
    bc(77) = "413111"
    bc(78) = "241112"
    bc(79) = "134111"
    bc(80) = "111242"
    bc(81) = "121142"
    bc(82) = "121241"
    bc(83) = "114212"
    bc(84) = "124112"
    bc(85) = "124211"
    bc(86) = "411212"
    bc(87) = "421112"
    bc(88) = "421211"
    bc(89) = "212141"
    bc(90) = "214121"
    bc(91) = "412121"
    bc(92) = "111143"
    bc(93) = "111341"
    bc(94) = "131141"
    bc(95) = "114113"
    bc(96) = "114311"
    bc(97) = "411113"
    bc(98) = "411311"
    bc(99) = "113141"
    bc(100) = "114131"
    bc(101) = "311141"
    bc(102) = "411131"
    bc(103) = "211412"
    bc(104) = "211214"
    
    bc(105) = "211232"   'START C
    bc(106) = "2331112"   'STOP
    '=======================
    
    '==================================
    If calculaDVCodBar128C(pSequenciaCodBar, varDV) = False Then
        pMensagemErro = "Não foi possível gerar o código de barras! Erro na geração do dígito verificador."
        DesenhaCodigoBarras128C = False
        Exit Function
    End If
    '==================================
                        
    '================ Cria padrão do Cód. Barras 128C ========
    
    'START C
    ReDim varPadraoCodBarras128C(0)
    varPadraoCodBarras128C(0) = bc(105)
    
    For i = 1 To Len(pSequenciaCodBar) Step 2
        
        'próximas 2 posições da sequência
        varPosicao = CInt(Mid(pSequenciaCodBar, i, 2))
        
        'Cria +1 posição no array
        ReDim Preserve varPadraoCodBarras128C(UBound(varPadraoCodBarras128C) + 1)
        
        'Registra padrão ref. número da sequência
        varPadraoCodBarras128C(UBound(varPadraoCodBarras128C)) = bc(varPosicao)
    
    Next
    
    'Cria +1 posição no array
    ReDim Preserve varPadraoCodBarras128C(UBound(varPadraoCodBarras128C) + 1)
    'Registra padrão ref. digito verificador
    varPadraoCodBarras128C(UBound(varPadraoCodBarras128C)) = bc(varDV)
    
    'STOP
    ReDim Preserve varPadraoCodBarras128C(UBound(varPadraoCodBarras128C) + 1)
    varPadraoCodBarras128C(UBound(varPadraoCodBarras128C)) = bc(106)
    
    '=========================================================
                    
    '=========== Dimensões ===================================
    pObjeto.ScaleMode = 3 'pixels
    
    pObjeto.Height = pHeightCodBarras
    pObjeto.FontSize = pTamanhoFonte
    
    pObjeto.Cls
    pObjeto.Picture = Nothing
    dw = CInt(pObjeto.ScaleHeight / 40) 'espaco entre as barras
    
    If dw < 1 Then dw = 1
        
    th = pObjeto.TextHeight(pSequenciaCodBar)   'altura do texto
    tw = pObjeto.TextWidth(pSequenciaCodBar)    ' largura do texto
    new_string = Chr$(1) & pSequenciaCodBar & Chr$(2)
    
    y1 = pObjeto.ScaleTop
    y2 = pObjeto.ScaleTop + pObjeto.ScaleHeight - IIf((pExibirNumeroCodBarras = True), (1.5 * th), 2)
    
    pObjeto.Width = 1.1 * Len(new_string) * (15 * dw) * pObjeto.Width / pObjeto.ScaleWidth
    '==========================================================
    
    'desenha cada caractere na string do codigo de barras
    xpos = pObjeto.ScaleLeft
            
    'Line(Flags As Integer, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, Color As Long)
    
    'Margem clara...5 espaços no início
    pObjeto.Line (xpos, y1)-(xpos + 5 * dw, y2), &HFFFFFF, BF
    xpos = xpos + (5 * dw)
            
    'Para cada padrão faça...Ex.:(0) = "211232" ... (1) = "2331112" ...
    For i = 0 To UBound(varPadraoCodBarras128C)
            
        'Combinação de barras: B = barra preta e S = espaço (barra branca)
        'CODE C = B S B S B S
        varDesenharBarra = True
        
        'Para cada item do padrão faça...Ex: padrão = 211232 ... (1)= 2 ... (2) = 1 ... (3) = 1 ... (4) = 2 ...
        For n = 1 To Len(varPadraoCodBarras128C(i))
            
            'próximo caracter do padrão...
            c = Mid(varPadraoCodBarras128C(i), n, 1)
            
            If varDesenharBarra = True Then
                'Desenha Barra Preta
                pObjeto.Line (xpos, y1)-(xpos + CInt(c) * dw, y2), &H0&, BF
                xpos = xpos + CInt(c) * dw
                                
                varDesenharBarra = False
            Else
                'Desenha Barra Branca
                pObjeto.Line (xpos, y1)-(xpos + CInt(c) * dw, y2), &HFFFFFF, BF
                xpos = xpos + CInt(c) * dw
                
                varDesenharBarra = True
            End If
            
        Next
    Next
        
    'Margem clara...5 espaços no final
    pObjeto.Line (xpos, y1)-(xpos + 5 * dw, y2), &HFFFFFF, BF
    xpos = xpos + (5 * dw)
           
    'tamanho final e texto
    pObjeto.Width = (xpos + dw) * pObjeto.Width / pObjeto.ScaleWidth
    pObjeto.CurrentX = (pObjeto.ScaleWidth - tw) / 2
    pObjeto.CurrentY = y2 + 0.25 * th

    If pExibirNumeroCodBarras = True Then
        pObjeto.Print pSequenciaCodBar
    End If
    
    'copia para o clipboard
    pObjeto.Picture = pObjeto.Image
    Clipboard.Clear
    Clipboard.SetData pObjeto.Image, 2
    DesenhaCodigoBarras128C = True
    Exit Function
    
trataErro:
            
    pMensagemErro = "Não foi possível gerar o código de barras! Erro: " & Err.Number & " - " & Err.Description
    DesenhaCodigoBarras128C = False
    
End Function

Public Function GetTempDir() As String
    
    Dim strFolder As String
    Dim lngResult As Long
    
    strFolder = String(MAX_PATH, 0)
    lngResult = GetTempPath(MAX_PATH, strFolder)
    
    If lngResult <> 0 Then
        If Right(Left(strFolder, lngResult), 1) = "\" Then
            GetTempDir = Left(strFolder, lngResult)
        Else
            GetTempDir = Left(strFolder, lngResult) & "\"
        End If
    Else
        GetTempDir = ""
    End If
    
End Function
