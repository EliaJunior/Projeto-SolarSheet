Attribute VB_Name = "a_functions"
Option Explicit
Option Base 0
Function ColunaValorUpdate(xForm As Object, ws As Worksheet) As Variant
    Dim c As Control, cTag As String, cNome As String, cTipo As String, cValor As Variant
    Dim par As New Scripting.Dictionary
    
    For Each c In xForm.Controls
        cTag = c.Tag
        cNome = c.name
        cTipo = TypeName(c)
        If InArray(Array("ComboBox", "TextBox", "OptionButton", "CheckBox"), cTipo) And ColunaExiste(cNome, ws, 1) Then
            cValor = c.value
            If StringInString(cTag, "campo-data") Then
                par.Add cNome, cNome & "= #" & cValor & "#"
            ElseIf StringInString(cTag, "numeric") Then
                par.Add cNome, cNome & "=" & cValor
            Else
                par.Add cNome, cNome & "='" & cValor & "'"
            End If
        End If
    Next c
    ColunaValorUpdate = par.Items
End Function
Sub ColunaValorInsertInto(xForm As Object, ws As Worksheet, colunas As Variant, valores As Variant, colIdNome As String, colStatusNome As String, curId As Long)
    Dim c As Control, cTag As String, cNome As String, cTipo As String, cValor As Variant
    Dim valor As New Scripting.Dictionary
    Dim cols As New Scripting.Dictionary
    
    For Each c In xForm.Controls
        cTag = c.Tag
        cNome = c.name
        cTipo = TypeName(c)
        If InArray(Array("ComboBox", "TextBox", "OptionButton", "CheckBox"), cTipo) And ColunaExiste(cNome, ws, 1) Then
            cValor = c.value
            If StringInString(cTag, "campo-data") Then
                valor.Add cNome, "#" & cValor & "#"
            ElseIf StringInString(cTag, "numeric") Then
                valor.Add cNome, cValor
            Else
                valor.Add cNome, "'" & cValor & "'"
            End If
            cols.Add cNome, cNome
        End If
    Next c
    'Adicona o ID no dicionario
    valor.Add colIdNome, curId
    cols.Add colIdNome, colIdNome
    
    'Adiciona o status no dicionario
    valor.Add colStatusNome, "'ativo'"
    cols.Add colStatusNome, colStatusNome
    
    colunas = cols.Items
    valores = Array(valor.Items)
End Sub
Function ThisWorkbookFullPath() As String
    Dim oneDrivePart As String
    Dim FullPath As String
    
    FullPath = ThisWorkbook.FullName
    FullPath = VBA.Replace(FullPath, "/", "\")
    oneDrivePart = "https:\\d.docs.live.net\"
    If VBA.InStr(FullPath, oneDrivePart) Then
        FullPath = VBA.Replace(FullPath, oneDrivePart, "")
        FullPath = Right(FullPath, Len(FullPath) - VBA.InStr(1, FullPath, "\"))
        FullPath = Environ$("OneDriveConsumer") & "\" & FullPath
    End If
    ThisWorkbookFullPath = FullPath
End Function
Function ThisWorkbookPath() As String
    Dim oneDrivePart As String
    Dim xPath As String
    
    xPath = ThisWorkbook.path
    xPath = VBA.Replace(xPath, "/", "\")
    oneDrivePart = "https:\\d.docs.live.net\"
    If VBA.InStr(xPath, oneDrivePart) Then
        xPath = VBA.Replace(xPath, oneDrivePart, "")
        xPath = Right(xPath, Len(xPath) - VBA.InStr(1, xPath, "\"))
        xPath = Environ$("OneDriveConsumer") & "\" & xPath
    End If
    ThisWorkbookPath = xPath
End Function
Function DadosListView(lvx As Object) As Variant
    Dim arOut As Variant
    Dim i As Long, j As Long
    Dim uc As Long, lC As Long
    Dim uR As Long, lR As Long
    Dim vlAux As Variant
    
    
    'Dimensão do array
    lC = 0
    lR = 0
    
    uc = lvx.ColumnHeaders.Count - 1
    uR = lvx.ListItems.Count
    
    ReDim arOut(lR To uR, lC To uc)
    
    'Capturando dados
    With lvx
        'Cabeçalho
        
        For i = lR To uR
            If i = lR Then
                For j = 0 To .ColumnHeaders.Count - 1
                    arOut(i, j) = .ColumnHeaders(j + 1).Text
                Next j
            ElseIf i > lR Then
                arOut(i, lC) = .ListItems(i)
                For j = 1 To .ColumnHeaders.Count - 1
                    vlAux = .ListItems(i).ListSubItems(j).Text
                    arOut(i, j) = vlAux
                Next j
            End If
        Next i
        
    End With
    
    DadosListView = arOut
    
End Function
Function PathExiste(path As String, isFile As Boolean) As Boolean
    Dim fso As New Scripting.FileSystemObject

    On Error Resume Next
        If isFile Then
            PathExiste = fso.FileExists(path)
        Else
            PathExiste = fso.FolderExists(path)
        End If
    On Error GoTo 0
    
    Set fso = Nothing
End Function
Function TransporMatriz(arIn As Variant) As Variant
    '------------------------------------------------------------------------------------
    'Retorna a matriz transposta de um array qualquer
    '   e.g. arIn(a to b, c to d) -> arOut(c to d, a to d)
    '   Par metros:
    '       arIn  -> Array de entrada
    '
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    ' ltima modifica  o 20/12/2023
    '------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim lR As Long, uR As Long, lC As Long, uc As Long
    Dim arOut() As Variant
    Dim arAux As Variant
    
    'Caso seja range, transforma em matriz
    arAux = arIn
    
    'verifica se   um array n o vazio
    If Not isArrayNotEmpty(arAux) Then
        TransporMatriz = arAux
    End If
    
    'Verifica se   uma array bidimensional
    If isArray2D(arAux) Then GoTo transpor2D Else GoTo transpor1D
    
transpor2D:
    lR = LBound(arAux)
    uR = UBound(arAux)
    lC = LBound(arAux, 2)
    uc = UBound(arAux, 2)
    
    ReDim arOut(lC To uc, lR To uR)
    For i = LBound(arAux) To UBound(arAux)
        For j = LBound(arAux, 2) To UBound(arAux, 2)
            arOut(j, i) = arAux(i, j)
        Next j
    Next i
    TransporMatriz = arOut
Exit Function
transpor1D:

    lR = LBound(arAux)
    uR = UBound(arAux)
    lC = LBound(arAux)
    uc = LBound(arAux)
    ReDim arOut(lC To uc, lR To uR)
    
    For i = LBound(arAux) To UBound(arAux)
        arOut(lC, i) = arAux(i)
    Next i

    TransporMatriz = arOut
End Function
Function isArray2D(arIn As Variant) As Boolean
    Dim lR As Long, lC As Long
    
    lR = LBound(arIn)
    
    On Error GoTo fim
    lC = LBound(arIn, 2)
    
    isArray2D = True
    Exit Function
fim:
    isArray2D = False
End Function
Function ListaMeses() As Variant
    Dim arOut(0 To 11, 0 To 1)
    Dim i As Long
    
    For i = 0 To 11
        arOut(i, 0) = UCase(Format("01/" & Format(i + 1, "00") & "/2000", "MMMM"))
        arOut(i, 1) = i + 1
    Next i
    
    ListaMeses = arOut
End Function
Function HexToIndexNumber(ByVal HexColor As String) As Long
    Dim R As Integer
    Dim G As Integer
    Dim B As Integer
    Dim cell As Range
  
    'Remover #
    HexColor = Replace(HexColor, "#", "")
    
    '6 caracteres
    HexColor = Right$("000000" & HexColor, 6)
    
    'Extrair valores RGB
    R = Val("&H" & Mid(HexColor, 1, 2))
    G = Val("&H" & Mid(HexColor, 3, 2))
    B = Val("&H" & Mid(HexColor, 5, 2))
    
    HexToIndexNumber = RGB(R, G, B)
End Function
Function HexToRGB(ByVal HexColor As String, r_g_b As String) As Long
    Dim R As Integer
    Dim G As Integer
    Dim B As Integer
    Dim cell As Range
  
    'Remover #
    HexColor = Replace(HexColor, "#", "")
    
    '6 caracteres
    HexColor = Right$("000000" & HexColor, 6)
    
    'Extrair valores RGB
    R = Val("&H" & Mid(HexColor, 1, 2))
    G = Val("&H" & Mid(HexColor, 3, 2))
    B = Val("&H" & Mid(HexColor, 5, 2))
    
    Select Case r_g_b
        Case "R"
            HexToRGB = R
        Case "G"
            HexToRGB = G
        Case "B"
            HexToRGB = B
        Case Else
            HexToRGB = R
    End Select
End Function
Function StringDelimitada(xString As String, xDelimitadorI As String, xDelimitadorF As String) As String
    Dim stOut As String
    Dim stLen As Long, cInicial As Long
    
    cInicial = InStr(xString, xDelimitadorI)
    stLen = Abs(InStr(xString, xDelimitadorI) - InStr(xString, xDelimitadorF) + 1)
    stOut = Mid(xString, cInicial + 1, stLen)
    StringDelimitada = stOut
End Function
Function TransporArray(arIn As Variant, arHeaderx As Variant)
    '------------------------------------------------------------------------------------
    'Retorna a matriz transposta de um array bidimensional de base 0
    '   Parâmetros:
    '       arIn        -> Array bidimensional de entrada
    '       arHeaderx   -> Array unidimensional de cabeçalho ou Empty/0 para desconsiderar
    '                   -> Deve conter a mesma quantidade de colunas do arIn
    '
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '------------------------------------------------------------------------------------
    
    Dim arOut() As Variant, dr As Long, dc As Long
    Dim i As Long, j As Long
    
    'Retorna matriz transposta
    If IsEmpty(arIn) And IsEmpty(arHeaderx) Then
        TransporArray = Empty
        Exit Function
    End If
    If IsEmpty(arIn) Then
        dr = 1
        dc = UBound(arHeaderx)
        ReDim arOut(0 To dr, 0 To dc)
        For i = LBound(arHeaderx) To UBound(arHeaderx)
            arOut(0, i) = arHeaderx(i)
        Next i
        TransporArray = arOut
        Exit Function
    End If
    
    'Dimensionando array de saída
    If Not IsEmpty(arHeaderx) Then
        dr = UBound(arIn, 2) + 1
    Else
        dr = UBound(arIn, 2)
    End If
    dc = UBound(arIn, 1)
    ReDim arOut(0 To dr, 0 To dc)
    For i = LBound(arIn) To UBound(arIn)
        If Not IsEmpty(arHeaderx) Then
            arOut(0, i) = arHeaderx(i)
            For j = LBound(arIn, 2) To UBound(arIn, 2)
                arOut(j + 1, i) = arIn(i, j)
            Next j
        Else
            For j = LBound(arIn, 2) To UBound(arIn, 2)
                arOut(j, i) = arIn(i, j)
            Next j
        End If
    Next i
    TransporArray = arOut
End Function
Function TabelaRefSQL(ws As Worksheet, RefTipo As Long, tbNome As String) As String
    '----------------------------------------------------------------------------------------------
    'Retorna o nome da tabela formatado para realização de consultas SQL
    'Exemplo: [TabelaX$A1:A2] AS [NomeTabela]
    '   Parâmetros:
    '       ws      -> Objeto Worksheet que contém o ListObject ou range
    '       RefTipo -> (0) Range | (1) ListObject | (2) Apenas Worksheet
    '       tbNome  -> Nome da tabela
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 05/06/2024
    '   Incluso RefTipo = 2
    '----------------------------------------------------------------------------------------------
    Dim rngEndereco As String
    Dim nomePl As String
    Dim ul As Long, uc As Long
    
    nomePl = ws.name & "$"
      
    With ws
        If RefTipo = 0 Then
            rngEndereco = .ListObjects(1).Range.Address(False, False)
        ElseIf RefTipo = 1 Then
            ul = .Cells(Rows.Count, 1).End(xlUp).Row
            uc = .Cells(1, Columns.Count).End(xlToLeft).Column
            rngEndereco = .Range(.Cells(1, 1), .Cells(ul, uc)).Address(False, False)
        ElseIf RefTipo = 2 Then
            TabelaRefSQL = "[" & nomePl & "] " & IIf(tbNome <> "", "AS [" & tbNome & "] ", "")
            Exit Function
        Else
            TabelaRefSQL = "[RefTipo] inválido!"
            Exit Function
        End If
    End With
    
    TabelaRefSQL = "[" & nomePl & rngEndereco & "] " & IIf(tbNome <> "", "AS [" & tbNome & "] ", "")
End Function
Function rowMatch(arIn1 As Variant, stProc As String, colProc As Long) As Long
    '------------------------------------------------------------------------------------
    'Retora o número da linha correspondente ao valor procurado, na coluna indicada
    '   Parâmetros:
    '       x
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '------------------------------------------------------------------------------------
    Dim i As Long
    For i = LBound(arIn1) To UBound(arIn1)
        If stProc = arIn1(i, colProc) Then
            rowMatch = i
            Exit Function
        Else
            rowMatch = -1
        End If
    Next i
End Function
Function TratarNulo(xVl As Variant, retornoSeVerdadeiro As Variant) As Variant
    '------------------------------------------------------------------------------------
    'Retora determinado valor se a entrada for vazio, e retorna o proprio valor caso contrário
    '   Parâmetros:
    '       xVl                 -> Valor de entrada
    '       retornoSeVerdadeiro -> Valor de saída caso seja nulo
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 11/02/2024
    '------------------------------------------------------------------------------------
    If IsNull(xVl) Then TratarNulo = retornoSeVerdadeiro Else TratarNulo = xVl
End Function

Function TratarVazios(arIn As Variant, VazioTratado As Variant) As Variant
    'IMPORTANTE'
    '------------------------------------------------------------------------------------
    'Retora algo
    '   Parâmetros:
    '       x
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim vlAux As Variant
    
    If IsEmpty(arIn) Then Exit Function
    For i = LBound(arIn) To UBound(arIn)
        For j = LBound(arIn, 2) To UBound(arIn, 2)
            vlAux = arIn(i, j)
            If IsEmpty(vlAux) Or IsNull(vlAux) Then
                arIn(i, j) = VazioTratado
            End If
        Next j
    Next i
    TratarVazios = arIn
End Function
Function FormatarNumerosArray(arIn As Variant, stFormato As String) As Variant
    '-------------------------------------------------------------------------------------
    'Retorna os valores numéricos do array de entrada de acordo com a string de formatação
    '   Parâmetros:
    '       x
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '-------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim vlAux As Variant
        For i = LBound(arIn) To UBound(arIn)
            For j = LBound(arIn, 2) To UBound(arIn, 2)
                vlAux = arIn(i, j)
                If IsNumeric(vlAux) Then
                    If IsEmpty(vlAux) Then arIn(i, j) = 0
                    arIn(i, j) = Format(arIn(i, j), xFormato)
                End If
            Next j
        Next i
    FormatarArray = arIn
End Function
Function TratarDouble(xVl As Variant) As Double
    'IMPORTANTE'
    '-------------------------------------------------------------------------------------
    'Retorna um valor qualquer para um valor numérico (forçadamente)
    '   Parâmetros:
    '       xvl -> Valor a ser tratado
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '-------------------------------------------------------------------------------------
    If Not IsNumeric(xVl) Then TratarDouble = 0 Else TratarDouble = CDbl(xVl)
End Function
Function TratarLong(xVl As Variant) As Long
    'IMPORTANTE'
    '-------------------------------------------------------------------------------------
    'Retorna um valor qualquer para um valor numérico (forçadamente)
    '   Parâmetros:
    '       xvl -> Valor a ser tratado
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '-------------------------------------------------------------------------------------
    If Not IsNumeric(xVl) Then TratarLong = 0 Else TratarLong = CLng(xVl)
End Function
Function DividirSafe(xNum As Variant, yDiv As Variant, Optional nDecimal As Variant) As Double
    '-------------------------------------------------------------------------------------
    'Retorna a divisão entre dois números, retornando 0 em caso de erro
    '   Parâmetros:
    '       xNum        -> Numerador
    '       yDiv        -> Divisor
    '       nDecimal    -> Numero de casas decimais
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 10/02/2024
    '-------------------------------------------------------------------------------------
    xNum = TratarDouble(xNum)
    yDiv = TratarDouble(yDiv)
    If yDiv = 0 Then
        DividirSafe = 0
    Else
        DividirSafe = xNum / yDiv
    End If
    
    If Not IsMissing(nDecimal) Then
        DividirSafe = Round(DividirSafe, CLng(nDecimal))
    End If
    
End Function
Function StringParaDate(xIn As Variant) As Date
    '-------------------------------------------------------------------------------------
    'Retorna a data de acordo com a string de entrada
    '   Parâmetros:
    '       x
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '-------------------------------------------------------------------------------------
    If IsDate(xIn) Then TransformarData = CDate(xIn) Else TransformarData = CDate(False)
End Function
Function CapturarColuna(arIn As Variant, nCol As Long, wthHeader As Boolean) As Variant
    '-------------------------------------------------------------------------------------
    'Captura determinada coluna de um array, retorna um array unidimensional
    '   Parâmetros:
    '       arIn        -> Array de dados
    '       nCol        -> Índice da coluna que se deseja capturar
    '       wthHeader   -> Se true, captura também o cabeçalho
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '-------------------------------------------------------------------------------------
    On Error GoTo saida
    Dim i As Long
    Dim drL As Long, drU As Long
    Dim arOut() As Variant
    
    If IsEmpty(arIn) Then
        CapturarColuna = arIn
        Exit Function
    End If
    
    drL = LBound(arIn) + IIf(wthHeader, 0, -1)
    drU = UBound(arIn)
    ReDim arOut(drL To drU)
    
    For i = LBound(arIn) + IIf(wthHeader, 0, 1) To UBound(arIn)
        arOut(i) = arIn(i, nCol)
    Next i
    CapturarColuna = arOut
    Exit Function
saida:
End Function
Function SomarColuna(arIn As Variant, xCol As Long) As Double
    '-------------------------------------------------------------------------------------
    'Retorna a soma dos elementos de determinada coluna de um array
    '   Parâmetros:
    '       x
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '-------------------------------------------------------------------------------------
    
    Dim i As Long
    Dim vlAux As Long
    If IsEmpty(arIn) Then
        SomarColuna = 0
        Exit Function
    End If
    For i = LBound(arIn) To UBound(arIn)
        vlAux = TratarDouble(arIn(i, xCol))
        SomarColuna = SomarColuna + vlAux
    Next i
End Function
Function InArray(arIn As Variant, vl As Variant) As Boolean
    '-------------------------------------------------------------------------------------
    'Verifica se determinado valor pertence ao array de entrada
    '   Parâmetros:
    '       x
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '-------------------------------------------------------------------------------------
    Dim i As Long
    InArray = False
    If IsArray(arIn) Then
        For i = LBound(arIn) To UBound(arIn)
            If arIn(i) = vl Then
                InArray = True
                Exit Function
            End If
        Next i
    Else
        If arIn = vl Then
            InArray = True
            Exit Function
        End If
    End If
End Function
Function ArrayAppend(arIn As Variant, arToAppend) As Variant
    '--------------------------------------------------------------------------------------
    'Adiciona uma nova linha no array
    '   Parâmetros:
    '       arIn        -> Array original
    '       arToAppend  -> Linha que será incluída no array original
    '   Observações
    '       (1) O array de entrada deve ter a mesma quantidade de colunas que o array a ser
    '           incluído, o array a ser incluído deve ter apenas uma linha
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 08/02/2024
    '--------------------------------------------------------------------------------------
    Dim arOut As Variant
    Dim lbl As Long, ubL As Long
    Dim lbC As Long, ubC As Long
    Dim lbLArToAppend As Long
    Dim i As Long, j As Long
    
    'Se o array de entrada for vazio, então apenas será igual ao array de inclusão
    If IsEmpty(arIn) Then
        arOut = arToAppend
        ArrayAppend = arOut
        Exit Function
    End If
    
    lbl = LBound(arIn, 1)                                                   'Linha inicial
    ubL = UBound(arIn, 1) + UBound(arToAppend) - LBound(arToAppend) + 1     'Linha final + novas linhas
    lbC = LBound(arIn, 2)                                                   'Coluna inicial
    ubC = UBound(arIn, 2)                                                   'Coluna final
    
    lbLArToAppend = LBound(arToAppend)
    
    'Redimensionando array
    ReDim arOut(lbl To ubL, lbC To ubC)
    
    'Copiando os dados originais
    For i = LBound(arIn, 1) To UBound(arIn, 1)
        For j = LBound(arIn, 2) To UBound(arIn, 2)
            arOut(i, j) = arIn(i, j)
        Next j
    Next i
    
    'Incluindo a nova linha
    For j = LBound(arIn, 2) To UBound(arIn, 2)
        For i = LBound(arToAppend) To UBound(arToAppend)
            arOut(UBound(arIn) + 1 + i, j) = arToAppend(i, j)
        Next i
    Next j
    ArrayAppend = arOut
End Function
Function BubbleSort(arIn As Variant) As Variant
    'Sorts a one-dimensional VBA array from smallest to largest
    'using the bubble sort algorithm.
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim arOut As Variant
    
    If IsEmpty(arIn) Then Exit Function
    
    arOut = arIn
     
    For i = LBound(arOut) To UBound(arOut) - 1
        For j = i + 1 To UBound(arOut)
            If arOut(i) > arOut(j) Then
                temp = arOut(j)
                arOut(j) = arOut(i)
                arOut(i) = temp
            End If
        Next j
    Next i
    
    BubbleSort = arOut
End Function

Function ArReplace(arIn As Variant, xTextoAntigo As String, xTextoNovo As String) As Variant
    '-------------------------------------------------------------------------------------
    'Substitui strings de um array bidimensional
    '   Parâmetros:
    '       x
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '-------------------------------------------------------------------------------------
    
    Dim i As Long, j As Long
    Dim arOut As Variant
    
    arOut = arIn
    For i = LBound(arIn) To UBound(arIn)
        For j = LBound(arIn, 2) To UBound(arIn, 2)
            arOut(i, j) = Replace(arIn(i, j), xTextoAntigo, xTextoNovo)
        Next j
    Next i
    ArReplace = arOut
End Function
Function ConsultaSQL(stSQL As String, withHeader As Boolean, hasHeader As Boolean, dataSource As String, isTextFile As Boolean) As Variant
    '---------------------------------------------------------------------------------------------
    'Retorna um array de acordo com string de consulta SQL
    '   Parâmetros:
    '       stSQL       ->  String de consulta SQL
    '       withHeader  ->  Retornar array com cabeçalho na primeira linha
    '       hasHeader   ->  Indica se a fonte de dados possui cabeçalho na primeira linha
    '       dataSource  ->  Caminho do local onde se encontram os dados
    '
    '
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 04/06/2024
    '---------------------------------------------------------------------------------------------
    
    Dim dbConn As New ADODB.Connection, rs As New ADODB.Recordset
    Dim arOut As Variant, headers As Variant, i As Long
    Dim stHDR As String
    Dim exProperties As String
    

    'String de conexão
    stHDR = IIf(hasHeader, "YES", "NO")
    
    If isTextFile Then
        exProperties = "'Text;HDR=" & stHDR & ";FMT=Delimited';"
    Else
        exProperties = "'Excel 12.0 Xml;HDR=" & stHDR & ";IMEX=1';"
    End If
    
    'Abre conexão
    With dbConn
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dataSource & ";Extended Properties=" & exProperties
        .Open
    End With
    
    'Abre recordset
    On Error Resume Next
        rs.Open stSQL, dbConn, adOpenKeyset, adLockOptimistic
        If Err.Number <> 0 Then
            ConsultaSQL = Err.Description
            Debug.Print Err.Description
            Exit Function
        End If
    On Error GoTo 0
    
    'Salva os dados em um array
    If Not rs.EOF Then arOut = rs.GetRows() Else arOut = Empty
    
    'Captura array e cabeçalho
    If withHeader = True Then
        With rs
            ReDim headers(0 To 0, 0 To .Fields.Count - 1)
            For i = 0 To .Fields.Count - 1
                headers(0, i) = .Fields(i).name
            Next i
        End With
        arOut = TransporMatriz(arOut)
        arOut = ArrayAppend(headers, arOut)
    Else
        If Not IsEmpty(arOut) Then
            arOut = TransporMatriz(arOut)
        End If
        headers = Empty
    End If
    
    If IsEmpty(arOut) And withHeader Then arOut = TransporMatriz(headers)
    
    'Desconectar bd
    rs.Close
    dbConn.Close
    Set rs = Nothing
    Set dbConn = Nothing
    
    'Retorno de função
    ConsultaSQL = arOut
End Function
Function ValidarFormulario(xFormulario As Object, arCampos As Variant) As Boolean
    '-------------------------------------------------------------------------------------
    'Retorna True se cada campo obrigatório de um formulário foi preenchido
    '   Parâmetros:
    '       xFormulario     -> UserForm a ser verificado
    '       arObrigatorio   -> Array(x,y,z) que contém os controles obrigatórios
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 22/12/2023
    'Data de modificação 13/04/2024
    '   Alterado variável para tipo Object
    '   Alterado lógica de validação
    '   Incluso verificação do tipo de campo de acordo com a tag
    '-------------------------------------------------------------------------------------
    Dim c As Control, curControl As String, curValue As Variant, curNomeCampo As String
    Dim cTag As String
    Dim i As Long
    
    On Error Resume Next
    ValidarFormulario = True
    For i = LBound(arCampos) To UBound(arCampos)
        Set c = arCampos(i)
        curValue = c.value
        cTag = c.Tag
        If (curValue = "") Or (StringInString(cTag, "numeric") And Not IsNumeric(curValue)) Or _
           (StringInString(cTag, "campo-data") And Not IsDate(curValue)) Then
            c.BackColor = RGB(248, 203, 173)
            ValidarFormulario = False
        Else
            c.BackColor = RGB(255, 255, 255)
        End If
    Next i

    If ValidarFormulario = False Then MsgBox "Os campos destacados são obrigatórios, confira o tipo de entrada!", vbExclamation
End Function
Function ValidarCamposNumericos(xFrame As Object, arCampos As Variant) As Boolean
    '-------------------------------------------------------------------------------------
    'Retorna True se os campos indicados forem numéricos
    '   Parâmetros:
    '       usfx            -> UserForm a ser verificado
    '       arCampos   -> Array(x,y,z) que contém o nome dos controles obrigatórios
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 12/03/2024
    '-------------------------------------------------------------------------------------
    Dim c As Control, curControl As String, curValue As Variant, curNomeCampo As String
    On Error Resume Next
    ValidarCamposNumericos = True
    For Each c In xFrame.Controls
        curControl = c.name
        If InArray(arCampos, curControl) Then
            curValue = c.value
            If Not IsNumeric(curValue) Then
                c.BackColor = RGB(248, 203, 173)
                ValidarCamposNumericos = False
            Else
                c.BackColor = RGB(255, 255, 255)
            End If
        End If
    Next c
    If ValidarCamposNumericos = False Then MsgBox "Os campos marcados devem ser valores numéricos!", vbExclamation
End Function
Function ValidarFrame(xFrame As Object, arObrigatorio As Variant) As Boolean
    'IMPORTANTE'
    '-------------------------------------------------------------------------------------
    'Retorna True se cada campo obrigatório de um frame foi preenchido
    '   Parâmetros:
    '       xFrame          -> Frame a ser verificado
    '       arObrigatorio   -> Array(x,y,z) que contém o nome dos controles obrigatórios
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 04/03/2024
    '-------------------------------------------------------------------------------------
    Dim c As Control, curControl As String, curValue As Variant, curNomeCampo As String
    On Error Resume Next
    ValidarFrame = True
    For Each c In xFrame.Controls
        curControl = c.name
        If InArray(arObrigatorio, curControl) Then
            curValue = c.value
            If c.value = "" Then
                c.BackColor = RGB(248, 203, 173)
                ValidarFrame = False
            Else
                c.BackColor = RGB(255, 255, 255)
            End If
        End If
    Next c
    If ValidarFrame = False Then MsgBox "Os campos marcados são obrigatórios!", vbExclamation
End Function

Function ExisteRegistros(xColuna As String, xValorProcurado As String, xTabela As String) As Boolean
    '--------------------------------------------------------------------------------------
    'Retorna TRUE caso existam registros na tabela e FALSE caso contrário
    '   Parâmetros:
    '       xColuna         -> Nome da coluna de referência
    '       xValorProcurado -> Valor procurado na coluna de referência
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 26/12/2023
    '--------------------------------------------------------------------------------------
    Dim SLC As Variant, frm As String, WHRE As Variant
    Dim cSQL As String
    Dim arOut As Variant

    SLC = Array(xColuna)
    WHRE = Array(SQLWhere(xColuna, xValorProcurado, False))
    cSQL = SQLQueryString(SLC, xTabela, WHRE)
    arOut = ConsultaSQL(cSQL, False, ThisWorkbook.FullName)
    If IsEmpty(arOut) Then
        ExisteRegistros = False
    Else
        ExisteRegistros = True
    End If
End Function
Function ColunaLBoxSelecionados(xLBox As Object, xCol As Long) As Variant
    '------------------------------------------------------------------------
    'Retorna as linhas selecionadas da coluna especificada de um ListBox
    '   Parâmetros:
    '       xLBox   -> ListBox
    '       xCol    -> Número da coluna
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 26/12/2023
    '------------------------------------------------------------------------
    Dim dicSelecao As New Scripting.Dictionary, x As Long

    With xLBox
        For x = 0 To .ListCount - 1
            If .Selected(x) Then
                dicSelecao.Add x, .List(x, xCol)
            End If
        Next x
    End With
    
    ColunaLBoxSelecionados = dicSelecao.Items
End Function
Function ColunaLViewSelecionados(xLView As Object) As Variant
    '------------------------------------------------------------------------
    'Retorna as linhas selecionadas da coluna especificada de um ListView
    '   Parâmetros:
    '       xLView   -> ListView
    '   Observações:
    '       (1) A coluna ID deve ser a primeira coluna do ListView
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 28/01/2024
    '------------------------------------------------------------------------
    Dim dicSelecao As New Scripting.Dictionary, x As Long
    Dim ct As Long
    
    With xLView
        For x = 1 To .ListItems.Count
            If .ListItems.Item(x).Selected Then
                dicSelecao.Add x, CLng(.ListItems(x).Text)
                ct = ct + 1
            End If
        Next x
    End With
    If ct = 0 Then
        ColunaLViewSelecionados = Array(-1)
    Else
        ColunaLViewSelecionados = dicSelecao.Items
    End If
End Function
Function ForcarMissing(Optional xValue As Variant, Optional xMiss) As Variant
    '----------------------------------------------------------------------------------------------
    'Retorna o valor da variável, se o valor não for passado, ou é vazio ou é 0, retorna 'missing'
    '   Parâmetros:
    '       xValue  -> Valor qualquer
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 29/12/2023
    '----------------------------------------------------------------------------------------------
    If xValue = "" Or IsMissing(xValue) Or (xValue = 0 And TypeName(xValue) <> "Boolean") Then
        ForcarMissing = xMiss
    Else
        ForcarMissing = xValue
    End If
End Function
Function ColIndex(colNome As String, arIn As Variant) As Long
    '----------------------------------------------------------------------------------------------
    'Retorna o índice da coluna de um array, de acordo com o nome da coluna
    '   Parâmetros:
    '       colNome -> Nome da coluna
    '       arIn    -> Array de entrada
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 29/12/2023
    '----------------------------------------------------------------------------------------------
    Dim i As Long, curCol As Variant
    
    'Coluna de um array bidimensinal (x to y, z to s)
    On Error Resume Next
    For i = LBound(arIn, 2) To UBound(arIn, 2)
        curCol = CStr(arIn(LBound(arIn, 1), i))
        If curCol = colNome Then
            ColIndex = i
            Exit Function
        End If
    Next i
    
    'Sai da função caso não haja erro, ou seja, o array era Bidimensional
    If Err.Number = 0 Then Exit Function
    
    'Coluna de um array Unidimensional (x to y)
    For i = LBound(arIn) To UBound(arIn)
        curCol = arIn(i)
        If curCol = colNome Then
            ColIndex = i
            Exit Function
        End If
    Next i
End Function
Function VerificarRegistro(vlRegistro As String, colNome As String, ws As Worksheet, Optional filtroStatusAtivo As Variant, Optional nomeColunaStatus As String) As Boolean
    '--------------------------------------------------------------------------------------------
    'Verifica se determinado registro existe em determinada coluna
    '   Parâmetros:
    '       vlRegistro -> Valor/String do registro que se deseja verificar
    '       colNome    -> Nome da coluna que se deseja verificar se o registro é existente
    '       plNome     -> Nome da tabela onde se encontram os dados
    '       Observações:
    '               (1) A planilha deve conter um ListObject(1)
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 29/01/2024
    '--------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim ySELECT As Variant, yWHERE As Variant
    Dim nomeTabela As String
    Dim arOut As Variant, cnt As Variant
    
    nomeTabela = TabelaRefSQL(ws, 0, "tDados")
    
    
    strSQL = "SELECT " & _
                "COUNT(*) " & _
            "FROM " & _
                "$tabela$ " & _
            "WHERE " & _
                "[tDados].[$colref$] = '$registro$' "
    
    If Not IsMissing(filtroStatusAtivo) Then strSQL = strSQL & " AND [tDados].[$colunastatus$] = 'ativo'"
    
    strSQL = Replace(strSQL, "$tabela$", nomeTabela)
    strSQL = Replace(strSQL, "$colref$", colNome)
    strSQL = Replace(strSQL, "$registro$", vlRegistro)
    strSQL = Replace(strSQL, "$colunastatus$", nomeColunaStatus)
 
    arOut = ConsultaSQL(strSQL, False)
    
    If isArrayNotEmpty(arOut) Then cnt = TratarLong(arOut(0, 0)) Else cnt = 0
    
    VerificarRegistro = IIf(cnt > 0, True, False)
End Function

Function CapturarArrayLV(lvx As Object, HeaderFRow As Boolean) As Variant
    '-------------------------------------------------------------------------------------
    'Retorna os dados de um ListView
    '   Parâmetros:
    '       lvx         -> ListView que será preenchido
    '       HeaderFRow  -> True caso a primeira linha seja de cabeçalho
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '-------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim ct As Long
    Dim Ldr As Long, Udr As Long, Ldc As Long, Udc As Long
    Dim lItem As ListItem
    Dim arOut As Variant
    Dim start As Long
    
    Ldr = 0
    Udr = lvx.ListItems.Count - 1 + IIf(HeaderFRow, 1, 0)   '+ 1 linha se tiver cabeçalho
    Ldc = 0
    Udc = lvx.ColumnHeaders.Count - 1
    
    ReDim arOut(Ldr To Udr, Ldc To Udc)
    If HeaderFRow Then start = Ldr Else start = Ldr + 1
    
    'Capturando dados
    With lvx
        For i = 1 To .ListItems.Count
            Set lItem = .ListItems(i)
            arOut(i + start, Ldc) = lItem.Text
            For j = 1 To lItem.ListSubItems.Count
                arOut(i + start, j) = lItem.ListSubItems(j)
            Next j
        Next i
    
    
    'Capturando cabeçalho se tiver
    If HeaderFRow Then
        For j = 1 To .ColumnHeaders.Count
            arOut(Ldr, j - 1) = .ColumnHeaders(j).Text
        Next j
    End If
    
    End With
    CapturarArrayLV = arOut
End Function
Function tbNovaLinha(lo As ListObject) As Long
    'IMPORTANTE'
    '-----------------------------------------------------------------------------
    'Adiciona nova linha no listobject e retorna o ID
    '   Parâmetros:
    '       lo  -> ListObject
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 04/02/2024
    '------------------------------------------------------------------------------
    Dim idOut As Long
    
    With lo
        idOut = .ListRows.Count + 1
        .ListRows.Add
    End With
    
    tbNovaLinha = idOut
End Function
Function ListViewSelectedID(lvx As Object) As Long
    '-----------------------------------------------------------------------------
    'Retorna o ID selecionado de uma ListView
    '   Parâmetros:
    '       lvx  -> ListView que contém a coluna ID
    '   Observações:
    '       (1) A coluna de ID deve ser a primeira coluna do LV
    '       (2) Retorna 0 caso não haja itens selecionados
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 05/03/2024
    '------------------------------------------------------------------------------
    Dim id As Long
    On Error GoTo retornarzero
    If Not lvx.SelectedItem.Selected Then GoTo retornarzero
        id = TratarLong(lvx.SelectedItem)
        ListViewSelectedID = id
    Exit Function
retornarzero:
    ListViewSelectedID = 0
End Function
Function ComboBoxSelectedID(cbx As Object, nCol As Long) As Long
    '-----------------------------------------------------------------------------
    'Retorna o ID selecionado de uma ComboBox, de acordo com a coluna indicada
    '   Parâmetros:
    '       cbx  -> ComboBox que contém a coluna ID
    '   Observações:
    '       (1) Retorna 0 caso não haja itens selecionados
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 06/03/2024
    '------------------------------------------------------------------------------

    ComboBoxSelectedID = 0
    If cbx.ListIndex = -1 Then Exit Function
    
    ComboBoxSelectedID = TratarLong(cbx.List(cbx.ListIndex, nCol))

End Function
Function ListViewIsSelected(lvx As Object) As Boolean
    On Error GoTo elsex
    If lvx.SelectedItem.Selected Then
        ListViewIsSelected = True
    Else
elsex:
         ListViewIsSelected = False
         MsgBox "Não há registros selecionados!", vbExclamation
    End If
End Function
Function ListBoxIsSelected(lvx As Object) As Boolean
    On Error GoTo elsex
    If lvx.ListIndex > 0 Then
        ListBoxIsSelected = True
    Else
elsex:
         ListBoxIsSelected = False
         MsgBox "Não há registros selecionados!", vbExclamation
    End If
End Function
Function ListBoxSelectedID(lvx As Object) As Long
    '-----------------------------------------------------------------------------
    'Retorna o ID selecionado de uma ListBox
    '   Parâmetros:
    '       lvx  -> ListBox que contém a coluna ID
    '   Observações:
    '       (1) A coluna de ID deve ser a primeira coluna do LV
    '       (2) Retorna 0 caso não haja itens selecionados
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 30/03/2024
    '------------------------------------------------------------------------------
    Dim id As Long, idx As Long
    On Error GoTo retornarzero
    If Not lvx.ListIndex <> -1 Then GoTo retornarzero
        idx = TratarLong(lvx.ListIndex)
        id = lvx.List(idx, 0)
        ListBoxSelectedID = id
    Exit Function
retornarzero:
    ListBoxSelectedID = 0
End Function

Function AreUSure(qst As String) As Boolean
    'Mensagem de confirmação de algo'
    
    If MsgBox(qst, vbExclamation + vbYesNo) = vbNo Then
        AreUSure = False
    Else
        AreUSure = True
    End If
End Function
Function isComboBoxSelected(cbx As Object, withalert As Boolean, Optional msgalert As Variant) As Boolean
    '-----------------------------------------------------------------------------------------
    'Verifica se a combobox possui seleção e retorna verdadeiro caso sim, falso caso contrário
    '   Parâmetros:
    '       cbx   -> Array que contém objetos Object
    '       xTipo   -> Tipo de entrada
    '   Observações:
    '       [1] Para valores numéricos
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 08/02/2024
    '-----------------------------------------------------------------------------------------
    Dim idx As Long
    Dim stAlert As String
    
    If Not IsMissing(msgalert) Then stAlert = msgalert Else stAlert = "Selecione um registro!"
    
    idx = cbx.ListIndex
    If idx = -1 Then
        isComboBoxSelected = False
        If withalert Then
            MsgBox stAlert, vbExclamation
        End If
    Else
        isComboBoxSelected = True
    End If
End Function
Function isTextBoxThisType(arTBX As Variant, xTipo As Long)
    '-----------------------------------------------------------------------------
    'Verifica se a propriedade value das TextBox contidas no array de entrada é de determinado tipo
    '   Parâmetros:
    '       arTBX   -> Array que contém objetos Object
    '       xTipo   -> Tipo de entrada
    '   Observações:
    '       [1] Para valores numéricos
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 08/02/2024
    '------------------------------------------------------------------------------
    Dim i As Long, vlAux As Variant
    Dim retorno As Boolean
    
    
    retorno = True
    
    'Iteração entre todas as tbox
    For i = LBound(arTBX) To UBound(arTBX)
        vlAux = arTBX(i).value
        Select Case xTipo
            Case 0
                If Not IsNumeric(vlAux) Then
                    retorno = False
                    Exit For
                End If
        End Select
    Next i
    
    isTextBoxThisType = retorno
End Function
Function MinimoEntre(arValores As Variant) As Double
    '-----------------------------------------------------------------------------
    'Retorna o menor valor contido no array
    '   Parâmetros:
    '       arValores   -> Array que contém objetos Object
    '   Observações:
    '       (1) Caso o valor não seja numérico, será considerado 0
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 08/02/2024
    '------------------------------------------------------------------------------
    Dim i As Long, vlAux As Double
    
    vlAux = TratarDouble(arValores(LBound(arValores)))
    
    For i = LBound(arValores) To UBound(arValores)
        If TratarDouble(arValores(i)) < vlAux Then
            vlAux = TratarDouble(arValores(i))
        End If
    Next i
    
    MinimoEntre = vlAux
End Function
Function MaximoEntre(arValores As Variant) As Double
    '-----------------------------------------------------------------------------
    'Retorna o maior valor contido no array
    '   Parâmetros:
    '       arValores   -> Array que contém objetos Object
    '   Observações:
    '       (1) Caso o valor não seja numérico, será considerado 0
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 08/02/2024
    '------------------------------------------------------------------------------
    Dim i As Long, vlAux As Double
    
    vlAux = TratarDouble(arValores(LBound(arValores)))
    
    For i = LBound(arValores) To UBound(arValores)
        If TratarDouble(arValores(i)) > vlAux Then
            vlAux = TratarDouble(arValores(i))
        End If
    Next i
    
    MaximoEntre = vlAux
End Function
Function MediaEntreValores(arValores As Variant, Optional RoundOf As Variant) As Double
    '-----------------------------------------------------------------------------
    'Retorna a média dos valores contidos no array
    '   Parâmetros:
    '       arValores   -> Array que contém objetos Object
    '   Observações:
    '       (1) Caso o valor não seja numérico, não será considerado na média
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 08/02/2024
    '------------------------------------------------------------------------------
    Dim i As Long, vlAux As Double
    Dim qtElementos As Long
    Dim casas As Long
    
    If Not IsMissing(RoundOf) Then
        casas = TratarLong(RoundOf)
    Else
        casas = 2
    End If
    
    If IsEmpty(arValores) Or Not IsArray(arValores) Then Exit Function
    
    On Error GoTo bidim
    
    'Quantidade de elementos
    For i = LBound(arValores) To UBound(arValores)
        If IsNumeric(arValores(i)) Then
            qtElementos = qtElementos + 1
        End If
    Next i
    
    For i = LBound(arValores) To UBound(arValores)
        If IsNumeric(arValores(i)) Then
            vlAux = TratarDouble(arValores(i)) / qtElementos + vlAux
            vlAux = Round(vlAux, casas)
        End If
    Next i
    
    MediaEntreValores = vlAux
    Exit Function
bidim:
    'Array bidirecional
    
    On Error GoTo finalizar
    'Quantidade de elementos
    For i = LBound(arValores) To UBound(arValores)
        If IsNumeric(arValores(i, LBound(arValores))) Then
            qtElementos = qtElementos + 1
        End If
    Next i
    
    For i = LBound(arValores) To UBound(arValores)
        If IsNumeric(arValores(i, LBound(arValores))) Then
            vlAux = TratarDouble(arValores(i, LBound(arValores))) / qtElementos + vlAux
            vlAux = Round(vlAux, casas)
        End If
    Next i
    
    MediaEntreValores = vlAux
    Exit Function
finalizar:
    MediaEntreValores = 0
End Function
Function ShapeExists(wsx As Worksheet, shName As String) As Boolean
    '-----------------------------------------------------------------------------
    'Retorna verdadeiro caso o shape exista na planilha indicada
    '   Parâmetros:
    '       wsx     -> Array Worksheet
    '       shName  -> Nome do shape procurado
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 20/02/2024
    '------------------------------------------------------------------------------
    Dim vlAux As String
    On Error Resume Next
        vlAux = wsx.Shapes(shName).name
        If Err.Number <> 0 Then ShapeExists = False Else ShapeExists = True
    On Error GoTo 0
End Function
Function gfSerieExists(wsx As Worksheet, shName As String, serieNome As String) As Boolean
    '-----------------------------------------------------------------------------
    'Retorna verdadeiro caso o shape exista na planilha indicada
    '   Parâmetros:
    '       wsx     -> Array Worksheet
    '       shName  -> Nome do shape procurado
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 20/02/2024
    '------------------------------------------------------------------------------
    Dim gf As Chart
    Dim seriesCount As Long
    Dim curSerie As String
    Dim i As Long
    
    Set gf = wsx.Shapes(shName).Chart
    
    'Retorno padrão da função = false
    gfSerieExists = False
    
    On Error Resume Next
        'Se a contagem de series do gráfico for 0 entao retorna falso
        seriesCount = gf.FullSeriesCollection.Count
        If seriesCount = 0 Then Exit Function
        
        'Verifica o nome de cada série do gráfico e retorna verdadeiro caso alguma coincida com a entrada
        For i = 1 To seriesCount
            curSerie = gf.FullSeriesCollection(i).name
            If curSerie = serieNome Then
                gfSerieExists = True
                Exit Function
            End If
        Next i
    On Error GoTo 0
End Function
Function PegarCaminho(MsoTipo As Long, titulo As String, SelecaoMultipla As Boolean, Optional arExtensoes As Variant) As Variant
    'IMPORTANTE'
    '-----------------------------------------------------------------------------
    'Retorna o caminho do arquivo ou pasta selecionada
    '   Parâmetros:
    '       tipoArquivo  -> Pasta ou Arquivo
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 21/02/2024
    '------------------------------------------------------------------------------
    Dim pathOut As String
    Dim i As Long
    Dim curDesc As String, curExt As String
    
    With Application.FileDialog(MsoTipo)
        .Title = titulo
        .AllowMultiSelect = SelecaoMultipla
        
        If Not IsMissing(arExtensoes) Then
            curDesc = arExtensoes(0) 'Descrição da extensão
            curExt = arExtensoes(1)  'Extensão
            With .Filters
                .Clear
                .Add curDesc, curExt, 1
            End With
        End If
        
        .Show
        If .SelectedItems.Count = 0 Then
            PegarCaminho = ""
            Exit Function
        End If
        If Not SelecaoMultipla Then
            PegarCaminho = .SelectedItems(1)
        Else
            PegarCaminho = .SelectedItems
        End If
    End With
End Function
Function RemoverNulos(rgIn As Variant)
    Dim i As Long
    Dim arOut As Variant
    Dim arAux As Variant
    Dim arInsert(0, 0) As Variant
    
    arAux = rgIn
    arOut = Empty
    
    For i = LBound(arAux) To UBound(arAux)
        If arAux(i, 1) <> 0 Then
            arInsert(0, 0) = arAux(i, 1)
            arOut = ArrayAppend(arOut, arInsert)
        End If
    Next i
    
    RemoverNulos = arOut
    
End Function
Function WsExiste(nameWS As String) As Boolean
    '-----------------------------------------------------------------------------
    'Verifica se a planilha informada existe e retorna verdadeiro caso exista
    '   Parâmetros:
    '       nameWS  -> Nome da Worksheet
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 25/02/2024
    '------------------------------------------------------------------------------
    Dim idxWS As Long
    On Error Resume Next
    idxWS = ThisWorkbook.Worksheets(nameWS).Index
    
    If Err.Number > 0 Then WsExiste = False Else WsExiste = True
End Function
Function isArrayNotEmpty(arIn As Variant) As Boolean
    '---------------------------------------------------------------------------------
    'Verifica se a variável de entrada é array e não é vazio e retorna true ou false
    '   Parâmetros:
    '       arIn  -> array de entrada
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 04/03/2024
    '----------------------------------------------------------------------------------
    If IsArray(arIn) And Not IsEmpty(arIn) Then isArrayNotEmpty = True Else isArrayNotEmpty = False
End Function
Function JoinArray(arIn As Variant, Delimitador As String, IncluirNull As Boolean, IterarLinha As Boolean, nIdx As Long) As String
    '---------------------------------------------------------------------------------
    'Retorna os valores de um array bidimensional concatenados
    '   Parâmetros:
    '       arIn        -> array de entrada
    '       Delimitador -> String de concatenação
    '       IncluirNull -> True para incluir valores nules ou False para não incluir
    '       IterarLinha -> True para concatenar valores de linha ou False para concatenar valores de coluna
    '       nIdx        -> Número da coluna ou linha que deseja percorrer
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 06/03/2024
    '----------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim dimensao As Long
    Dim curValor As Variant
    Dim stOut As String
    
    If Not isArrayNotEmpty(arIn) Then
        JoinArray = ""
        Exit Function
    End If
    
    dimensao = IIf(IterarLinha, 2, 1)
    
    For i = LBound(arIn, dimensao) To UBound(arIn, dimensao)
        If IterarLinha Then curValor = arIn(nIdx, i) Else curValor = arIn(i, nIdx)
        If IsNull(curValor) Or curValor = "" Then
            If IncluirNull Then
                stOut = curValor & Delimitador & stOut
            End If
        Else
            stOut = curValor & Delimitador & stOut
        End If
    Next i
    If stOut <> "" Then
        stOut = left(stOut, Len(stOut) - 1)
    End If
    JoinArray = stOut
End Function
Function StringInString(stPrincipal As String, stContida As String) As Boolean
    '-----------------------------------------------------------------------------------------
    'Verifica se determinada cadeia de caracteres está contida em outra cadeia de caracteres
    '   Parâmetros:
    '       stPrincipal -> String que contém a stContida
    '       stContida   -> String que está contida na stPrincipal
    '
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 24/03/2024
    '--------------------------------------------------------------------------------------------
    If InStr(1, stPrincipal, stContida) > 0 Then StringInString = True Else StringInString = False
End Function
Function ConsultarRegistro(ws As Worksheet, colRetorno As String, colMatch As String, MatchRegistro As Variant, vlNulo As Variant) As Variant
    '----------------------------------------------------------
    'Consulta ID de determinada tabela, contida em uma worksheet
    '   Par metros:
    '       ws          -> Planilha onde est  contida a tabela
    '       colProcura  -> Nome da Coluna de retorno
    '       colID       -> Nome da coluna que cont m o ID
    '       idRegistro  -> Id do registro
    '       vlNulo      -> Valor caso n o haja registros
    '
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    ' ltima altera  o 22/03/2024
    '----------------------------------------------------------
    Dim vlOut As Variant
    Dim stSQL As String
    Dim tb As String
    Dim filtro As String
    
    If TypeName(MatchRegistro) = "String" Then
        filtro = "[tb].[" & colMatch & "] = '" & MatchRegistro & "'"
    Else
        filtro = "[tb].[" & colMatch & "] = " & MatchRegistro
    End If
    
    tb = TabelaRefSQL(ws, 0, "tb") & " "
    
    stSQL = "SELECT " & _
                "[tb].[" & colRetorno & "] " & _
            "FROM " & _
                tb & _
            "WHERE " & _
                filtro
                
    vlOut = ConsultaSQL(stSQL, False)
    If isArrayNotEmpty(vlOut) Then ConsultarRegistro = vlOut(0, 0) Else ConsultarRegistro = vlNulo
End Function
Function ColumnListObjectExists(lo As ListObject, columnName As String) As Boolean
    On Error Resume Next
    Dim idx As Long
    idx = lo.ListColumns(columnName).Index
    
    If Err.Number <> 0 Then ColumnListObjectExists = False Else ColumnListObjectExists = True
End Function
Function ColunaExiste(colNome As String, ByVal ws As Worksheet, nLinhaCabecalho As Long) As Boolean
    Dim curCol As String
    Dim j As Long
    Dim uc As Long
    
    With ws
        uc = .Cells(nLinhaCabecalho, Columns.Count).End(xlToLeft).Column
        For j = 1 To uc
            curCol = .Cells(nLinhaCabecalho, j).value
            If curCol = colNome Then
                ColunaExiste = True
                Exit Function
            End If
        Next j
    End With
    ColunaExiste = False
End Function

