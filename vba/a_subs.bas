Attribute VB_Name = "a_subs"
Option Explicit
Option Base 0
Sub PreencherDadosFormularioSELECT(xForm As MSForms.UserForm, id As Long, colIdNome As String, ws As Worksheet)

    Dim c As Control, cNome As String, cTipo As String, cTag As String
    Dim cValor As Variant
    Dim slcCol As New Scripting.Dictionary
    Dim slct As String
    Dim arOut As Variant
    Dim stSQL As String
    Dim tb As String
    Dim i As Long
    Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
    Dim stCon As String
    Dim fld As Field
    
    
    'Tabela
    tb = TabelaRefSQL(ws, 2, "t")
    
    'Colunas (SELECT)
    For Each c In xForm.Controls
        cNome = c.name
        cTipo = TypeName(c)
        cTag = c.Tag
        If InArray(Array("ComboBox", "TextBox", "OptionButton", "CheckBox"), cTipo) And ColunaExiste(cNome, ws, 1) Then
            slcCol.Add cNome, cNome
        End If
    Next c
    
    'Consulta SQL
    slct = Join(slcCol.Items, ",") & " "

    stSQL = "SELECT " & slct & _
            "FROM " & tb & _
            "WHERE " & _
                colIdNome & "=" & id
    
    'Abrindo conexão
    'Abre conexão
    With cn
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                            "Data Source=" & ThisWorkbookFullPath & ";" & _
                            "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';"
        .Open
    End With
    
    'Abre recordset
    rs.Open stSQL, cn, adOpenKeyset, adLockOptimistic
    
    'Preenche os campos
    For Each fld In rs.Fields
        cNome = fld.name
        cTag = xForm.Controls(cNome).Tag
        cValor = fld.value
            
        If StringInString(cTag, "percentual") Then
            cValor = TratarDouble(cValor) * 100
        ElseIf StringInString(cTag, "campo-data") Then
            cValor = CDate(cValor)
        End If
        
        xForm.Controls(cNome).value = cValor
        
    Next fld
    
    'Fechando conexão
    rs.Close
    cn.Close
    Set cn = Nothing
    Set rs = Nothing

End Sub
Sub InsertIntoTable(xColumns As Variant, xValues As Variant, ws As Worksheet, _
                    hasHeader As Boolean, ByVal dSource As String)
    '-------------------------------------------------------------------------------------
    'Insere novos dados na tabela
    '   Parâmetros:
    '       xColumns    -> Array com o nome das colunas a serem inclusos: Array([Col1], [Col2], [Col3],...)
    '       xValues     -> Array com o valor das colunas a serem inclusas:
    '                       Array(Array([Val1], [Val2], [Val3],...),
    '                             Array([Val4], [Val5], [Val6],...),
    '                             Array([Val7], [Val8], [Val9],...)...)
    '                      Obs: Caso os valores sejam do tipo string, é necessário uar aspas simples
    '       ws          -> Planilha onde estão localizado os dados
    '       hasHeader   -> Define se a primeira linha é um cabeçalho
    '       dSource     -> Caminho do arquivo de dados
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 05/06/2024
    '-------------------------------------------------------------------------------------
    
    Dim stSQL As String, stSqlAux As String
    Dim cn As New ADODB.Connection
    Dim stCon As String, stHDR As String
    Dim exProperties As String
    Dim tName As String, stValues As String, stColumns As String
    Dim i As Long

    'Alias da tabela
    tName = TabelaRefSQL(ws, 2, "")
    
    'String Colunas
    stColumns = "[" & Join(xColumns, "],[") & "]"
    
    'String valores
    'stValues = "(" & Join(xValues, ",") & ")"
    
    'String de conexão
    stHDR = IIf(hasHeader, "YES", "NO")
    
    'Propriedades
    exProperties = "'Excel 12.0 Xml;HDR=" & stHDR & "'"
    
    stCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dSource & ";Extended Properties=" & exProperties & ";"
    
    'String SQL
    stSQL = "INSERT INTO " & tName & _
                "(" & stColumns & ") " & _
                "VALUES $values$ "
    
    With cn
        .ConnectionString = stCon
        .Open
        For i = LBound(xValues) To UBound(xValues)
            stSqlAux = Replace(stSQL, "$values$", "(" & Join(xValues(i), ",") & ")")
            .Execute stSqlAux
        Next i
    End With
    
    'Fechando instâncias
    cn.Close
    Set cn = Nothing
End Sub
Sub UpdateTable(arKeys As Variant, KeysColName As String, _
                colValues As Variant, ws As Worksheet, _
                hasHeader As Boolean, ByVal dSource As String)
    '-------------------------------------------------------------------------------------
    'Atualiza determinada tabela de acordo com o seu(s) id(s)
    '   Parâmetros:
    '       arKeys      -> Array de Ids(chaves): arId = Array(x,y,z,..)
    '       KeysColName -> Nome da coluna Id (Chave)
    '       colValues   -> Pares Coluna/Valor: Array(Coluna1 = valor1, Coluna2 = valor2, ...)
    '                      Obs: Caso os valores sejam do tipo string, é necessário uar aspas simples
    '       ws          -> Planilha onde estão localizado os dados
    '       hasHeader   -> Define se a primeira linha é um cabeçalho
    '       dSource     -> Caminho do arquivo de dados
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 05/06/2024
    '-------------------------------------------------------------------------------------
    
    Dim stSQL As String
    Dim cn As New ADODB.Connection
    Dim stCon As String, stHDR As String
    Dim exProperties As String
    Dim tName As String, stCondition As String, stValues As String
    
    'Alias da tabela
    tName = TabelaRefSQL(ws, 1, "tDados")
    
    'Condição de busca
    stCondition = "[tDados].[" & KeysColName & "] IN (" & Join(arKeys, ",") & ")"
    
    'Pares Coluna/Valor
    stValues = Join(colValues, ",")
    
    'String de conexão
    stHDR = IIf(hasHeader, "YES", "NO")
    
    'Propriedades
    exProperties = "'Excel 12.0 Xml;HDR=" & stHDR & "'"
    
    stCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dSource & ";Extended Properties=" & exProperties & ";"
    
    'String SQL
    stSQL = "UPDATE " & tName & _
            "SET " & stValues & " " & _
            "WHERE " & stCondition
    
    'Estabelecendo conexão
    With cn
        .ConnectionString = stCon
        .Open
        .Execute stSQL
    End With

    'Fechando instâncias
    cn.Close
    Set cn = Nothing

End Sub
Sub PreencherTabela(Tabela As Object, arDados As Variant)
    Dim tbNome As String
    tbNome = TypeName(Tabela)
    
    If tbNome = "ListBox" Then
        PreencherListBox Tabela, arDados
        AjustarListBox Tabela, -0.1
    Else
        PreencherListView Tabela, arDados, True
    End If
End Sub
Sub FormatarFormulario(xFormulario As Object)
    Dim c As Control, cTipo As String
    Dim CorBorda As Long, corFundoTitulo As Long, corFundoFormulario As Long
    Dim stTAG As String
    Dim stTIP As String
    
    CorBorda = 2709764
    corFundoFormulario = RGB(197, 197, 197)
    
    For Each c In xFormulario.Controls
        cTipo = TypeName(c)
        stTAG = c.Tag
        With c
            'Tags
            If StringInString(stTAG, "notShow") Then .Visible = False
            If StringInString(stTAG, "Locked") Then .Locked = True
            If StringInString(stTAG, "NotEnabled") Then .Enabled = False
            If StringInString(stTAG, "LockedNotEnabled") Then
                .Enabled = False
                .Locked = True
            End If
            Select Case cTipo
                Case "Label"
                    .Font.Size = 11
                    .Font.name = "Roboto"
                    .ForeColor = RGB(0, 0, 0)
                    .BackStyle = fmBackStyleTransparent
                    
                    If StringInString(stTAG, "titulo-formulario") Then
                        .Font.Size = 16
                        .TextAlign = fmTextAlignCenter
                        .BackStyle = fmBackStyleTransparent
                        '.BorderColor = corFundoTitulo
                        '.BorderStyle = fmBorderStyleSingle
                        '.BorderColor = CorBorda
                    End If
            
                Case "ComboBox"
                    .Font.Size = 11
                    .Font.name = "Roboto"
                    .ShowDropButtonWhen = fmShowDropButtonWhenFocus
                    .BorderStyle = fmBorderStyleSingle
                    .BorderColor = CorBorda
                    .Height = 18
                    
                Case "TextBox"
                    .Font.Size = 11
                    .Font.name = "Roboto"
                    .BorderStyle = fmBorderStyleSingle
                    .BorderColor = CorBorda
                    .Height = 18
    
                Case "CommandButton"
                    .BackColor = RGB(255, 255, 255)
                    
                Case "OptionButton"
                    .Font.Size = 11
                    .Font.name = "Roboto"
                    .ForeColor = RGB(0, 0, 0)
                    .BackStyle = fmBackStyleTransparent
                    
                Case "Frame"
                    .BackColor = corFundoFormulario
                    .BorderStyle = fmBorderStyleSingle
                    .Caption = ""
                    .BorderColor = CorBorda
            End Select
        End With
    Next c
    
    'Formatando frame
    With xFormulario
        .BackColor = corFundoFormulario
        .BorderStyle = fmBorderStyleSingle
        .Caption = ""
        .BorderColor = CorBorda
    End With
End Sub
Sub LimparCampos(xUSF As Object)
    '-------------------------------------------------------------------------------------
    'Limpa todos os campos de um formulário, salvo aqueles com tag 'NaoApagar'
    '   Parâmetros:
    '       xUSF -> UserForm que contém os controles
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 07/03/2024
    '-------------------------------------------------------------------------------------
    Dim c As Control
    Dim cTipo As String

    For Each c In xUSF.Controls
        cTipo = TypeName(c)
        If (cTipo = "ComboBox" Or cTipo = "TextBox") And c.Tag <> "NaoApagar" Then
            c.value = ""
        End If
    Next c
End Sub

Sub OtimizarIniciar()
        With Application
                .Calculation = xlCalculationManual
                .DisplayAlerts = False
                .ScreenUpdating = False
                .EnableEvents = False
                .DisplayStatusBar = False
        End With
End Sub
Sub OtimizarFinalizar()
        With Application
                .Calculation = xlCalculationAutomatic
                .DisplayAlerts = True
                .ScreenUpdating = True
                .EnableEvents = True
                .DisplayStatusBar = True
        End With
End Sub
Sub AjustarListBoxesUSF(uf1 As Object)
    '-------------------------------------------------------------------------------------
    'Ajusta todas as ListBoxes de um UserForm
    '   Parâmetros:
    '       uf1 -> UserForm que contém as ListBoxes
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '-------------------------------------------------------------------------------------
    Dim stWidths As String
    Dim arWidths As Variant
    Dim vl As Long, curVl As Long
    Dim i As Long, j As Long
    Dim arLista2 As Variant
    Dim larguraMedia As Double, LarguraExtra As Double
    Dim CaixaLista As Control
    Dim altCM As Double

    LarguraExtra = 0.2
    
    On Error Resume Next
    'Percorre todas as listbox da userform
    For Each CaixaLista In uf1.Controls
        If TypeOf CaixaLista Is Object  Then
            arLista2 = CaixaLista.List
            larguraMedia = CaixaLista.Font.Size * 0.0352778 'Altura da fonte em cm
            If IsEmpty(arLista2) = False And arLista2 <> "" And arLista2 <> Null Then
                ReDim arWidths(LBound(arLista2, 2) To UBound(arLista2, 2))
                'Percorre todas as colunas captura o comprimento da maior string e salva no array de widths
                For j = LBound(arLista2, 2) To UBound(arLista2, 2)
                    vl = 0
                    For i = LBound(arLista2) To UBound(arLista2)
                        curVl = Len(arLista2(i, j))
                        If curVl > vl Then
                            vl = curVl
                        End If
                    Next i
                    If vl = 0 Then vl = 10
                    arWidths(j) = (vl * larguraMedia) + LarguraExtra & "cm"
                Next j
                stWidths = Join(arWidths, ";")
                CaixaLista.ColumnWidths = stWidths
            End If
        End If
    Next CaixaLista
End Sub
Sub AjustarListBox(xLBox As Object, LarguraExtra As Double)
    '-------------------------------------------------------------------------------------
    'Ajusta as colunas da ListBox
    '   Parâmetros:
    '       xLBox           -> ListBox
    '       LarguraExtra    -> Largura de acréscimo na propriedade Width da coluna
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 26/12/2023
    '-------------------------------------------------------------------------------------
    Dim stWidths As String
    Dim arWidths As Variant
    Dim vl As Long, curVl As Long
    Dim i As Long, j As Long
    Dim arLista2 As Variant
    Dim larguraMedia As Double
    
        arLista2 = xLBox.List
        larguraMedia = xLBox.Font.Size * 0.0352778    'Altura da fonte em cm
        If Not IsEmpty(arLista2) Then
            ReDim arWidths(LBound(arLista2, 2) To UBound(arLista2, 2))
            'Percorre todas as colunas captura o comprimento da maior string e salva no array de widths
            For j = LBound(arLista2, 2) To UBound(arLista2, 2)
                vl = 0
                For i = LBound(arLista2) To UBound(arLista2)
                    curVl = Len(arLista2(i, j))
                    If curVl > vl Then
                        vl = curVl
                    End If
                Next i
                If vl = 0 Then vl = 10
                arWidths(j) = (vl * larguraMedia) + LarguraExtra & "cm"
            Next j
            stWidths = Join(arWidths, ";")
            xLBox.ColumnWidths = stWidths
        End If
End Sub
Sub PreencherListView(lvx As Object, arIn As Variant, HeaderFRow As Boolean)
    'IMPORTANTE'
    '-------------------------------------------------------------------------------------
    'Preenche um ListView com os valores de um array
    '   Parâmetros:
    '       lvx         -> ListView que será preenchido
    '       arIn        -> Array que contém os dados
    '       HeaderFRow  -> True caso a primeira linha seja de cabeçalho
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '-------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim ct As Long
    Dim startRow As Long
    Dim Item As ListItem
    
    With lvx
        .Font.name = "Courier New"
        .Font.Size = 11
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .MultiSelect = True
        .Gridlines = True
        .Appearance = ccFlat
    End With
    
    If HeaderFRow Then
        startRow = LBound(arIn) + 1
        For j = LBound(arIn, 2) To UBound(arIn, 2)
            ct = ct + 1
            lvx.ColumnHeaders.Add ct, , arIn(LBound(arIn), j)
            If j = LBound(arIn, 2) Then lvx.ColumnHeaders(ct).Width = 50
        Next j
    Else
        startRow = LBound(arIn)
    End If
    
    For i = startRow To UBound(arIn)
        Set Item = lvx.ListItems.Add
        For j = LBound(arIn, 2) To UBound(arIn, 2)
            If j = LBound(arIn, 2) Then
                Item.Text = arIn(i, j)
            Else
                If Not IsNull(arIn(i, j)) Then
                Item.ListSubItems.Add , , arIn(i, j)
                Else
                Item.ListSubItems.Add , , ""
                End If
            End If
        Next j
    Next i
End Sub
Sub AjustarColunasLV(lvx As Object)
    '-------------------------------------------------------------------------------------
    'Executa algo
    '   Parâmetros:
    '       x
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '-------------------------------------------------------------------------------------
    Dim col As ColumnHeader
    Dim i As Long, maxLen As Long
    Dim larguraMedia As Double

    larguraMedia = 8
    
    For Each col In lvx.ColumnHeaders
        maxLen = 0
        ' Encontra o tamanho máximo da string na coluna atual
        For i = 1 To lvx.ListItems.Count
            If Len(lvx.ListItems(i).ListSubItems(col.Index).Text) > maxLen Then
                maxLen = Len(lvx.ListItems(i).ListSubItems(col.Index).Text)
            End If
        Next i
        
        ' Ajusta a largura da coluna para o tamanho máximo encontrado
        If maxLen > 0 Then
            col.Width = maxLen * larguraMedia
        End If
    Next col
End Sub
Sub LigarTelaCheia(rgScreen As Range, ws As Worksheet)
    '-------------------------------------------------------------------------------------
    'Executa algo
    '   Parâmetros:
    '       x
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '-------------------------------------------------------------------------------------
    With ws
        .Protect
        .Activate
        .ScrollArea = rgScreen.Address
    End With
    
    rgScreen.Select
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)" 'Oculta todas as guias de menu
        .DisplayAlerts = False
        .DisplayStatusBar = False
        .DisplayFullScreen = True
        .DisplayFormulaBar = False
    End With
    
        With ActiveWindow
        .DisplayHeadings = False  'Oculta os títulos de linha e coluna
        .DisplayHorizontalScrollBar = False 'Ocultar barra horizontal
        .DisplayVerticalScrollBar = False  'Ocultar barra vertical
        .DisplayWorkbookTabs = False
        .Zoom = True 'Aplica o zoom automático
        .Zoom = 100
        .SmallScroll UP:=Rows.Count
        End With
    
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
End Sub
Sub DesligarTelaCheia()
Attribute DesligarTelaCheia.VB_ProcData.VB_Invoke_Func = "O\n14"
    '-------------------------------------------------------------------------------------
    'Executa algo
    '   Parâmetros:
    '       x
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '-------------------------------------------------------------------------------------
    With Application
            .DisplayFullScreen = False
            .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)" 'Reexibe os menus
            .DisplayFormulaBar = True                           'Reexibir a barra de fórmulas
            .DisplayStatusBar = True                            'Reexibir a barra de status, disposta ao final da planilha
    End With
    With ActiveWindow
            .DisplayHeadings = True                             'Reexibir o cabeçalho da Pasta de trabalho
            .DisplayHorizontalScrollBar = True                  'Reexibir barra horizontal
            .DisplayVerticalScrollBar = True                    'Reexibir barra vertical
            .DisplayWorkbookTabs = True                         'Reexibir guias das planilhas
            .DisplayHeadings = True                             'Reexibir os títulos de linha e coluna
    End With
End Sub
Sub imprimirRelatorio(rg As Range, FullNamePath As String, modPaisagem As Boolean, pgsAltura As Long, pgsLargura As Long, abrirDepois As Boolean)
    '-------------------------------------------------------------------------------------
    'Imprime o range inputado
    '   Parâmetros:
    '       rg              -> Range a ser impresso
    '       FullNamePath    -> Caminho completo do arquivo X:\pasta1\...\NomeArquivo.pdf
    '       modPaisagem     -> True para imprimir em modo paisagem
    '       pgsAltura       -> Paginas de altura
    '       pgsLargura      -> Paginas de largura
    '       abrirDepois     -> True para vizualizar arquivo após imprimir
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 02/02/2024
    '-------------------------------------------------------------------------------------
    
    Dim Orientacao As Long
    Dim wsRG As Worksheet

    'Planilha do RG
    Set wsRG = ThisWorkbook.Worksheets(rg.Parent.name)
    
    Orientacao = IIf(modPaisagem, xlLandscape, xlPortrait)
    
    'Configuraçãoes de pagina
    With wsRG.PageSetup
        .Zoom = False
        .FitToPagesWide = pgsLargura
        .FitToPagesTall = IIf(pgsAltura = 0, 1, pgsAltura)
        .Orientation = Orientacao
        .CenterHorizontally = True
        .TopMargin = 2
        .BottomMargin = 1.5
        .LeftMargin = 1.5
        .RightMargin = 1.5
        .PaperSize = xlPaperA4
    End With
    
    ' Exporta para PDF
    rg.ExportAsFixedFormat Type:=xlTypePDF, fileName:=FullNamePath, Quality:=xlQualityStandard, _
                                     IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=abrirDepois
End Sub
Sub excluirSheet(plName As String)
    '-------------------------------------------------------------------------------------
    'Executa algo
    '   Parâmetros:
    '       plName  -> Exclui planilha (plName) desta pasta de trabalho
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 21/02/2024
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(plName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub
Sub InserirImagemEmRange(ByVal ws As Worksheet, ByVal rng As Range, ByVal imagePath As String, Optional imgNome As Variant)
    '-------------------------------------------------------------------------------------
    'Executa algo
    '   Parâmetros:
    '       x
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 20/12/2023
    '-------------------------------------------------------------------------------------
    Dim pic As Shape
    On Error Resume Next
    Set pic = ws.Shapes.AddPicture(fileName:=imagePath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                    left:=rng.left, top:=rng.top, Width:=-1, Height:=-1)
    With pic
        .LockAspectRatio = msoTrue  ' Mantém a proporção da imagem
        .Width = rng.Width          ' Ajusta a largura para ser igual à largura do intervalo
        .Height = rng.Height        ' Ajusta a altura para ser igual à altura do intervalo
        .Placement = xlMoveAndSize  ' Move e dimensiona a imagem com a célula
    End With
    If Not IsMissing(imgNome) Then imgNome = pic.name
    On Error GoTo 0
End Sub
Sub ShowCalPt(xTbox As Object, Optional ByVal titulo As String)
    'Chama o userform 'Calendário' e imprime o valor no textbox
    Dim myDate As Date

    If titulo = "" Then
        titulo = "Selecione a data"
    End If
    
    CalendarForm.Caption = titulo
    myDate = CalendarForm.GetDate(Language:="pt", FirstDayOfWeek:=Monday, SaturdayFontColor:=RGB(250, 0, 0), SundayFontColor:=RGB(250, 0, 0))
    If myDate > 0 Then xTbox.value = myDate
End Sub
Sub ExibirFrames(arFramesVisiveis As Variant, FrameExcept As Variant, usf As Object, _
                xTop As Double, xHeight As Double, xLeft As Double)
    '------------------------------------------------------------------------------------------------
    'Esconte todos os frames do UserForm, exceto os informados
    '   Parâmetros:
    '       arFramesVisiveis    -> Array(x,y,z) com o nome dos frames que serão exibidos
    '       FrameExcept         -> Array(s,t,u) com o nome dos frames que não sofrerão nenhum efeito
    '       xTop                -> Propriedade Top do Frame que será exibido
    '       xHeight             -> Propriedade Height do Frame que será exibido
    '       xLeft               -> Propriedade Left do Frame que será exibido
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 24/03/202
    '   24/03/2024 -> Incluído propriedade left
    '------------------------------------------------------------------------------------------------
    Dim c As Control
    Dim cNome As String, cTipo As String
    Dim frmNome As String
    
    'Frames do menu
    For Each c In usf.Controls
        cNome = c.name
        cTipo = TypeName(c)
        If cTipo = "Frame" Then
            If InArray(arFramesVisiveis, cNome) Then
                c.Visible = True
                c.top = xTop
                c.left = xLeft
                c.Height = xHeight
            ElseIf Not InArray(FrameExcept, cNome) Then
                c.Visible = False
                c.top = 1000
                c.Height = 0
            End If
        End If
    Next c
End Sub
Private Sub OrdenarListView(lvx As Object, clmHeader As Object)
    With lvx
        .SortKey = clmHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        .Sorted = True
    End With
End Sub
Sub FormatarTextBoxTelefone(tbx As Object)
    '-------------------------------------------------------------------------------------
    'Formata TextBox como número de telefone (DD) 9XXXX-XXXX
    '   Parâmetros:
    '       tbx -> TextBox a ser formatada
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 22/12/2023
    '-------------------------------------------------------------------------------------
    
    Dim phoneNumber As String
    Dim formattedNumber As String
    Dim i As Integer

    
    ' Remove caracteres não numéricos do número de telefone
    phoneNumber = tbx.Text
    For i = 1 To Len(phoneNumber)
        If IsNumeric(Mid(phoneNumber, i, 1)) Then
            formattedNumber = formattedNumber & Mid(phoneNumber, i, 1)
        End If
    Next i
    
    ' Verifica se o número tem pelo menos um dígito
    If Len(formattedNumber) > 0 Then
        ' Formata o número para o padrão (00) 9 0000-0000
        formattedNumber = "(" & left(formattedNumber, 2) & ") " & _
                           Mid(formattedNumber, 3, 1) & " " & _
                           Mid(formattedNumber, 4, 4) & "-" & _
                           Mid(formattedNumber, 8, Len(formattedNumber))
                           
        ' Define o texto formatado na TextBox
        tbx.Text = formattedNumber
    End If
End Sub
Sub FormatarTextBoxCPF_CNPJ(tbx As Object)
    '-------------------------------------------------------------------------------------
    'Formata TextBox como CPF ou CNPJ dependendo do tamanho da string
    '   Parâmetros:
    '       tbx -> TextBox a ser formatada
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 22/12/2023
    '-------------------------------------------------------------------------------------
    Dim inputx As String
    Dim formattedValue As String
    
    ' Remove caracteres não numéricos
    inputx = Replace(tbx.Text, ".", "")
    inputx = Replace(inputx, "/", "")
    inputx = Replace(inputx, "-", "")
    
    ' Verifica se é um CPF (11 dígitos) ou CNPJ (14 dígitos)
    If Len(inputx) = 11 Then
        ' Formata como CPF (000.000.000-00)
        formattedValue = left(inputx, 3) & "." & Mid(inputx, 4, 3) & "." & Mid(inputx, 7, 3) & "-" & Right(inputx, 2)
    ElseIf Len(inputx) = 14 Then
        ' Formata como CNPJ (00.000.000/0000-00)
        formattedValue = left(inputx, 2) & "." & Mid(inputx, 3, 3) & "." & Mid(inputx, 6, 3) & "/" & Mid(inputx, 9, 4) & "-" & Right(inputx, 2)
    ElseIf Len(inputx) > 14 Then
        ' Caso contrário, mantém o texto como está
        formattedValue = left(tbx.Text, 18)
    Else
        ' Caso contrário, mantém o texto como está
        formattedValue = tbx.Text
    End If
    
    ' Define o texto formatado na TextBox
    tbx.Text = formattedValue
    ' Mantém o cursor no final do texto
    tbx.SelStart = Len(tbx.Text)
End Sub
Sub UCTextBox(tbx As Object)
    ' Verifica se a caixa de texto não está vazia
    If Not tbx Is Nothing Then
        ' Formata o texto para maiúsculas
        tbx.value = UCase(tbx.value)
    End If
End Sub
Sub SalvarDadosFormulario(xFormulario As Object, WSData As Worksheet, xID As Long, NomeColID As String)
    '-------------------------------------------------------------------------------------------------------------------
    'Salva os campos de determinado formulário no arOut e em seguida preenche a linha da tabela com os dados capturados
    '   Parâmetros:
    '       xFormulario -> Frame ou UserForm que contém os campos (TextBox, ComboBox, etc)
    '       loData      -> ListObject (tabela de dados)
    '       xID         -> Número da linha onde os dados serão salvos
    '   Observações:
    '       (1) O nome dos controles devem ser iguais aos nomes das respectivas colunas
    '       (2) A tabela deve conter uma coluna ID e uma coluna status (ativo/inativo)
    '       (3) Todos os campos da tabela, com exceção da coluna status e id, devem existir no formulário
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 27/03/2024
    '
    'Atualizações
    '(27/03/2024)
    '   -> Procedimento melhorado generalizando controle Frame e Userform
    '   -> Código otimizado para permitir formulários que não contenham todas as colunas da tabela
    '   -> Legibilidade melhorada
    '
    '-------------------------------------------------------------------------------------------------------------------
    Dim c As Control, cTipo As String, cNome As String, cTag As String, cValor As Variant
    Dim loData As ListObject
    Dim arOut As Variant, lR As Long
    Dim idxColuna As Long, idxID As Long, idxStatus As Long
    
    Set loData = WSData.ListObjects(1)
    
    'Dados atuais da linha da tabela
    arOut = loData.ListRows(xID).Range.value
    lR = LBound(arOut)
    
    With loData
        idxID = .ListColumns(NomeColID).Index
        idxStatus = .ListColumns("status").Index
    End With
    
    'Capturando os dados do formulário
    For Each c In xFormulario.Controls
        cTipo = TypeName(c)
        cNome = c.name
        cTag = c.Tag
        If InArray(Array("ComboBox", "TextBox", "CheckBox", "OptionButton"), cTipo) Then
            If ColumnListObjectExists(loData, cNome) Then
                idxColuna = loData.ListColumns(cNome).Index
                cValor = c.value
                
                'Tratando valor
                If StringInString(cTag, "date") And IsDate(cValor) Then
                    cValor = CDate(cValor)
                ElseIf StringInString(cTag, "percentual") Then
                    cValor = TratarDouble(cValor) / 100
                ElseIf StringInString(cTag, "numeric") Then
                    cValor = TratarDouble(cValor)
                End If
                
                'Salvando no array
                arOut(lR, idxColuna) = cValor
            End If
        End If
    Next c
    
    'Campo status e campo id
    arOut(lR, idxID) = xID
    arOut(lR, idxStatus) = "ativo"
    
    'Preenchendo linha da tabela
    Application.ScreenUpdating = False
    loData.DataBodyRange(xID, 1).Resize(1, UBound(arOut, 2)).value = arOut
    Application.ScreenUpdating = True
End Sub
Sub SalvarDadosFrame(xFrame As Object, plDados As Worksheet, xID As Long, NomeColID As String)
    'IMPORTANTE'
    '-------------------------------------------------------------------------------------------------------------------
    'Salva os campos de determinado frame no arOut e em seguida preenche a linha da tabela com os dados capturados
    '   Parâmetros:
    '       xFrame    -> Frame que contém os campos (TextBox, ComboBox, etc)
    '       plDados -> WorkSheet que contém o ListObject (tabela de dados)
    '       xID     -> Número do ID onde os dados serão salvos
    '   Observações:
    '       (1) A planilha deve conter apenas um ListObject -> índice 1
    '       (2) O nome dos controles devem ser iguais aos nomes das respectivas colunas
    '       (3) A tabela deve conter uma coluna ID e uma coluna status (ativo/inativo)
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 04/03/2024
    '-------------------------------------------------------------------------------------------------------------------
    Application.ScreenUpdating = False
    Dim c As Control, cTipo As String, cNome As String, cValor As Variant
    Dim loData As ListObject
    Dim arControles As Variant
    Dim arColID() As Variant, arValues() As Variant
    Dim arOut() As Variant, colIDX As Long
    Dim ct As Long, j As Long
    
    arControles = Array("ComboBox", "TextBox", "OptionButton", "CheckBox")
    
    Set loData = plDados.ListObjects(1)
    ct = -1
    For Each c In xFrame.Controls
        cTipo = TypeName(c)
        cNome = c.name
        If InArray(arControles, cTipo) Then
            colIDX = loData.ListColumns(cNome).Index
            cValor = c.value
            
            'Formata valor como data, caso a tag do botão seja 'campo-data'
            If c.Tag = "campo-data" And cValor <> "" Then
               cValor = CDate(cValor)
            ElseIf c.Tag = "percentual" Then
                cValor = TratarDouble(cValor) / 100
            End If
            
            ct = ct + 1
            ReDim Preserve arValues(0 To ct)
            ReDim Preserve arColID(0 To ct)
            arValues(ct) = cValor
            arColID(ct) = colIDX
        End If
    Next c
    
    'Preenchendo array de saída com os valores capturados
    ReDim arOut(0 To 0, 0 To UBound(arValues) + 2) '+2 Referente às colunas status e ID
    
    For j = LBound(arValues) To UBound(arValues)
        arOut(0, arColID(j) - 1) = arValues(j)
    Next j
    
    'Preenchendo Colunas STATUS e ID
    With loData
        arOut(0, .ListColumns(NomeColID).Index - 1) = xID
        arOut(0, .ListColumns("status").Index - 1) = "ativo"
    End With
    
    'Preenchendo linha da tabela
    loData.DataBodyRange(xID, 1).Resize(1, UBound(arOut, 2) + 1).value = arOut
    Application.ScreenUpdating = True
End Sub
Sub InativarLinhasTabela(arIds As Variant, ws As Worksheet)
    'IMPORTANTE'
    '------------------------------------------------------------------------
    'Inativa os IDs de determinada tabela
    '   Parâmetros:
    '       arIDs   -> Array que contém os IDs a serem inativados
    '       plNome  -> Nome da planilha
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 06/03/2024
    '------------------------------------------------------------------------
    Dim lo As ListObject
    Dim i As Long
    
    If IsEmpty(arIds) Then Exit Sub
    
    Set lo = ws.ListObjects(1)
    
    With lo.ListColumns("status")
        For i = LBound(arIds) To UBound(arIds)
            .DataBodyRange(arIds(i)).value = "inativo"
        Next i
    End With
    
End Sub
Sub PreencherListBox(xLBox As Object, arLista As Variant)
    '-------------------------------------------------------------------------------------
    'Preenche a ListBox com os dados do array
    '   Parâmetros:
    '       xLBox   -> ListBox a ser preenchida
    '       arLista -> Array que contém os dados de preenchimento (se vazio exit function)
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última modificação 26/12/2023
    '-------------------------------------------------------------------------------------
    If IsEmpty(arLista) Then
        xLBox.List = Array("Sem registros")
        Exit Sub
    End If
    
    With xLBox
        .BorderStyle = fmBorderStyleSingle
        .Font.Size = 12
        .Font.name = "Roboto"
        .MultiSelect = fmMultiSelectExtended
        .List = arLista
        .ColumnCount = UBound(arLista, 2) + 1
    End With
End Sub
Sub PreencherDadosFormulario(xFormulario As Object, loData As ListObject, xID As Long)
    '-------------------------------------------------------------------------------------------------------
    'Preenche os campos do Frame/Userform com os valores da ListObject
    '   Parâmetros:
    '       xFormulario     ->  Formulario (Frame ou UserForm) que contém os controles a serem preenchidos
    '       loData          ->  ListObject [Workbooks(x).Worksheets(y).ListObject(z)]
    '       xID             ->  Número da linha em que os dados se encontram
    '       Observações:
    '           (1) As colunas do ListObject devem coincidir com o nome do campo do userform
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217||
    'Última modificação 27/03/2024
    '
    'Atualizações
    '(27/03/2024)
    '   -> Generalizado userforms e frames
    '(03/04/2024)
    '   -> Tratamento de erro para controles que não existam na tabela de dados incluso
    '-------------------------------------------------------------------------------------------------------
    Dim c As Control, cTipo As String, cNome As String, cTag As String, cValor As Variant
    Dim btTipo As Variant
    
    btTipo = Array("TextBox", "OptionButton", "ComboBox", "CheckBox")
    
    For Each c In xFormulario.Controls
        cTipo = TypeName(c)
        cNome = c.name
        cTag = c.Tag
        If ColumnListObjectExists(loData, cNome) And InArray(btTipo, cTipo) Then
            cValor = loData.ListColumns(cNome).DataBodyRange(xID).value
            
            If StringInString(cTag, "percentual") Then
                cValor = TratarDouble(cValor) * 100
            ElseIf StringInString(cTag, "campo-data") Then
                cValor = CDate(cValor)
            End If
            
            c.value = cValor
        End If
    Next c
End Sub
Sub PreencherDadosFrame(xFrame As Object, lo As ListObject, xID As Long)
    'IMPORTANTE'
    '---------------------------------------------------------------------------------------------
    'Preenche os campos do Frame com os valores da ListObject
    '   Parâmetros:
    '       xFrame   ->  Frame (UserForm)
    '       lo      ->  ListObject [Workbooks(x).Worksheets(y).ListObject(z)]
    '       xID     ->  Número da linha em que os dados se encontram
    '       Observações:
    '           (1) As colunas do ListObject devem coincidir com o nome do campo do userform
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217||
    'Última modificação 05/03/2024
    '---------------------------------------------------------------------------------------------
    Dim c As Control, cTipo As String, cNome As String
    Dim btTipo As Variant
    
    btTipo = Array("TextBox", "OptionButton", "ComboBox", "CheckBox")
    
    For Each c In xFrame.Controls
        cTipo = TypeName(c)
        cNome = c.name
        If InArray(btTipo, cTipo) Then
            If c.Tag = "percentual" Then
                c.value = lo.ListColumns(cNome).DataBodyRange(xID).value * 100
            ElseIf c.Tag = "campo-data" Then
                c.value = CDate(lo.ListColumns(cNome).DataBodyRange(xID).value)
            Else
                c.value = lo.ListColumns(cNome).DataBodyRange(xID).value
            End If
        End If
    Next c
End Sub
Sub DesativarIDBancoDados(lo As ListObject, arIds As Variant)
    '---------------------------------------------------------------------------------------------
    'Preenche os campos do formulário com os valores da ListObject
    '   Parâmetros:
    '       lo      ->  ListObject [Workbooks(x).Worksheets(y).ListObject(z)]
    '       arIDS   ->  Array dos IDS que serão inativados
    '       Observações:
    '           (1) As colunas da ListBox devem coincidir com o nome do campo do userform
    '           (2) Deve existir uma coluna chamada 'status' na listObject
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217||
    'Última modificação 27/01/2024
    '---------------------------------------------------------------------------------------------
    Dim i As Long, curId As Long
    
    For i = LBound(arIds) To UBound(arIds)
        curId = CLng(arIds(i))
        lo.ListColumns("status").DataBodyRange(curId).value = "inativo"
    Next i
End Sub
Sub PreencherTextBoxComID(cbx As Object, tbxID As Object, vlNulo As Variant)
    '---------------------------------------------------------------------------------------------
    'Preenche uma TextBox com seu ID, capturado de determinada ComboBox
    '   Parâmetros:
    '       cbx     ->  ComboBox que contém duas colunas (valor | ID)
    '       tbxID   ->  TextBox que deverá ser preenchida com o ID
    '       vlNulo  ->  Valor de retorno em caso de seleção inválida
    '       Observações:
    '           (1) A ComboBox deve conter necessariamente 2 colunas
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217||
    'Última modificação 29/01/2024
    '---------------------------------------------------------------------------------------------
    Dim id As Variant
    If cbx.ListIndex = -1 Then
        id = 0
    Else
        id = cbx.List(cbx.ListIndex, 1)
    End If
    tbxID.value = id
End Sub
Sub FormatarTextBoxValor(tbx As Object)
    '---------------------------------------------------------------------------------------------
    'Formata a entrada de uma TextBox com o formato '0.0,00'
    '   Parâmetros:
    '       tbx  ->  TextBox de entrada
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217||
    'Última modificação 29/01/2024
    '---------------------------------------------------------------------------------------------
    Dim valor As String
    Dim formattedValue As String
    Dim decimalSeparator As String
    
    ' Define o separador decimal (pode variar dependendo das configurações do sistema)
    decimalSeparator = Application.International(xlDecimalSeparator)
    
    ' Remove caracteres não numéricos, exceto o separador decimal
    valor = Replace(tbx.Text, ",", "")
    valor = Replace(valor, ".", "")
    valor = Replace(valor, decimalSeparator, "")
    
    ' Verifica se há pelo menos um dígito
    If IsNumeric(valor) Then
        ' Formata o valor para o padrão 0,0.00
        formattedValue = Format(Val(valor) / 100, "0,0.00")
        
        ' Define o texto formatado na TextBox
        tbx.Text = formattedValue
    Else
        ' Se não for numérico, mantém o texto como está
        tbx.Text = 0
    End If
End Sub
Sub FormatarLinhaLVX(lvx As Object, nLinha As Long, corTexto As Long)
    Dim lvItem As ListItem, lvSubitem As ListSubItem
    
    Set lvItem = lvx.ListItems(nLinha)
    
    
    With lvItem
        .Bold = True
        .ForeColor = corTexto
        For Each lvSubitem In .ListSubItems
            lvSubitem.Bold = True
            lvSubitem.ForeColor = corTexto
        Next lvSubitem
    End With
End Sub
Sub PreencherGrafico(arValores As Variant, arRotulos As Variant, xGrafico As Chart, serieNome As Variant)
    '---------------------------------------------------------------------------------------------
    'Preenche os dados de determinada série de um gráfico
    '   Parâmetros:
    '       arValores       -> Array que contém os valores do gráfico
    '       arRotulos       -> Array que contém os rótulos do gráfico
    '       xGrafico        -> Gráfico que será preenchido 'Planilha.Shapes("x").chart'
    '       arSerieNome     -> Nome das série de dados contida no gráfico
    '   Observações:
    '       (1) Trata arrays vazios para retornar um gráfico vazio
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217||
    'Última modificação 14/02/2024
    '---------------------------------------------------------------------------------------------
    
    Dim arValAux As Variant, arRotuloAux As Variant

    'Trata os arrays caso vazios
    arValAux = TratarArrayVazio(arValores)
    arRotuloAux = TratarArrayVazio(arRotulos)

    With xGrafico.FullSeriesCollection(serieNome)
        .values = arValAux
        .xValues = arRotuloAux
    End With
End Sub
Sub excluirLinhaListObject(lo As ListObject, id As Long)
    '-----------------------------------------------------------------------------
    'Exclui linha do ListObject de arcodo com o ID
    '   Parâmetros:
    '       lo  -> ListObject
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 04/02/2024
    '------------------------------------------------------------------------------
    With lo
        .ListRows(id).Delete
    End With
End Sub
Sub BloquearDesbloquearCampos(arControls As Variant, liberar As Boolean)
    '-----------------------------------------------------------------------------
    'Bloqueia campos de uma userform
    '   Parâmetros:
    '       arControls  -> Array que contém os campos a serem bloqueados ou desbloqueados
    '       liberar     -> True para desbloquear, false para bloquear
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 08/02/2024
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim tipo As String
    
    For i = LBound(arControls) To UBound(arControls)
        tipo = TypeName(arControls(i))
        If tipo = "Label" Then
            arControls(i).Enabled = liberar
        Else
            arControls(i).Enabled = liberar
            arControls(i).Locked = Not liberar
        End If
    Next i
End Sub

Sub WriteAllBas()
' Write all VBA modules as .bas files to the directory of ThisWorkbook.
' Implemented to make version control work smoothly for identifying changes.
' Designed to be called every time this workbook is saved,
'   if code has changed, then will show up as a diff
'   if code has not changed, then file will be same (no diff) with new date.
' Following https://stackoverflow.com/questions/55956116/mass-importing-modules-references-in-vba
'            which references https://www.rondebruin.nl/win/s9/win002.htm

Dim cmp As VBComponent, cmo As CodeModule
Dim fn As Integer, outName As String
Dim sLine As String, nLine As Long
Dim dirExport As String, outExt As String
Dim fileExport As String

   On Error GoTo MustTrustVBAProject
   Set cmp = ThisWorkbook.VBProject.VBComponents(1)
   On Error GoTo 0
   dirExport = ThisWorkbook.path + Application.PathSeparator + "vba" + Application.PathSeparator
   For Each cmp In ThisWorkbook.VBProject.VBComponents
      Select Case cmp.Type
         Case vbext_ct_ClassModule:
            outExt = ".cls"
         Case vbext_ct_MSForm
            outExt = ".frm"
         Case vbext_ct_StdModule
            outExt = ".bas"
         Case vbext_ct_Document
            Set cmo = cmp.CodeModule
            If Not cmo Is Nothing Then
               If cmo.CountOfLines = cmo.CountOfDeclarationLines Then ' Ordinary worksheet or Workbook, no code
                  outExt = ""
               Else ' It's a Worksheet or Workbook but has code, export it
                  outExt = ".cls"
               End If
            End If ' cmo Is Nothing
         Case Else
            Stop ' Debug it
      End Select
      If outExt <> "" Then
         fileExport = dirExport + cmp.name + outExt
         If Dir(fileExport) <> "" Then Kill fileExport   ' From Office 365, Export method does not overwrite existing file
         cmp.Export fileExport
      End If
   Next cmp
   Debug.Print "Procedimento concluído! Backup concluído com sucesso!"
   Exit Sub
    
MustTrustVBAProject:
   MsgBox "Must trust VB Project in Options, Trust Center, Trust Center Settings ...", vbCritical + vbOKOnly, "WriteAllBas"
End Sub
Sub AjustarTabIndex(usfX As Object)
    '-----------------------------------------------------------------------------
    'Percorre todos os controles de uma userform e ajusta o TabIndex de cada um,
    'a ordem considerada é de cima para baixo, da esquerda para a direita
    '   Parâmetros:
    '       arControls  -> Array que contém os campos a serem bloqueados ou desbloqueados
    '       liberar     -> True para desbloquear, false para bloquear
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 08/02/2024
    '------------------------------------------------------------------------------
    Dim c As Control
    Dim arButtons As Variant
    Dim coordenadaX As Double, coordenadaY As Double, btNome As String, btTipo As String
    Dim TopDic As New Scripting.Dictionary
    Dim arTopFaixas As Variant
    Dim faixa As Double, curFaixa As Double
    Dim arCoordenadas As Variant
    Dim arAux(0 To 0, 0 To 2) As Variant
    Dim arClassificado As Variant
    Dim ct As Long
    Dim i As Long, j As Long
    Dim wsTemp As Worksheet, strSQL As String, tb As String
    
    arButtons = Array("CommandButton", "OptionButton", "CheckBox", "ComboBox", "TextBox")
    
    arCoordenadas = Empty
    
    'Captura o nome, top e left dos controles
    For Each c In usfX.Controls
        btTipo = TypeName(c)
        btNome = c.name
        coordenadaX = c.left
        coordenadaY = c.top
        arAux(0, 0) = btNome
        arAux(0, 1) = CDbl(coordenadaX)
        arAux(0, 2) = CDbl(coordenadaY)
        
        If InArray(arButtons, btTipo) Then
            If Not TopDic.Exists(coordenadaY) Then
                TopDic.Add coordenadaY, coordenadaY
            End If
            arCoordenadas = ArrayAppend(arCoordenadas, arAux)
        End If
    Next c
    
    'Cria uma tabela temporária para classificar o array usando SQL

    Application.ScreenUpdating = False
    excluirSheet "tempsheet"
    Set wsTemp = ThisWorkbook.Worksheets.Add
    wsTemp.name = "tempsheet"
    wsTemp.Range("A1").Resize(UBound(arCoordenadas) + 1, UBound(arCoordenadas, 2) + 1).value = arCoordenadas
    
    tb = RangeSQLRef("tbTemp", wsTemp.name)
    
    strSQL = "SELECT * FROM " & tb & " " & _
             "ORDER BY [tbTemp].[F3], [tbTemp].[F2] ASC"
             
    arClassificado = ConsultaSQL(strSQL, True, False)
    
    'Exclui planilha temporaria
    excluirSheet "tempsheet"
    Application.ScreenUpdating = True
    
    'Tab Indices
    For i = LBound(arClassificado) + 1 To UBound(arClassificado)
        usfX.Controls(arClassificado(i, 0)).TabIndex = i
    Next i
    
End Sub
Sub InserirGrafico(ws As Worksheet, gfStyle As Long, gfType As Long, gfTop As Double, _
                   gfLeft As Double, gfRotulos As Variant, gfValores As Variant, gfTitulo As String, _
                   gfWidth As Double, gfHeight As Double, gfBarShape As Long, gfNome As String, gfNomeSerie As String)
    '-----------------------------------------------------------------------------
    'Insere um gráfico na planilha de acordo com o tipo, valores, top e left
    '   Parâmetros:
    '       [1] ws              -> Worksheet onde o gráfico será inserido
    '       [2] gfStyle         -> Número correspondente ao estilo
    '                               [201] Coluna 2D
    '                               [286] Coluna 3D
    '       [3] gfType          -> Número correspondente ao tipo de gráfico
    '       [4] gfTop           -> Propriedade top do gráfico
    '       [5] gfLeft          -> Propriedade left do gráfico
    '       [6] gfWidth          -> Largura do gráfico
    '       [7] gfHeight         -> Altura do gráfico
    '       [8] gfRotulos       -> Rotulos do gráfico (XValues)
    '       [9] gfValores       -> Valores do gráfico (Values)
    '       [10] gfTitulo       -> Titulo do gráfico
    '       [11] gfBarShape     -> Formato da barra do gráfico
    '       [12] gfNome         -> Nome do gráfico
    '       [13] gfSerieNome    -> Nome da série
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 14/02/2024
    '------------------------------------------------------------------------------
    
    'Variaveis de procedimento
    Dim gf As Chart
    Dim grafExist As Boolean
    Dim ctSeries As Long

    'Verifica se o gráfico existe
    grafExist = ShapeExists(ws, gfNome)
    
    If Not grafExist Then 'Se o gráfico não existe então cria um novo e renomeia com o nome passado
        Set gf = ws.Shapes.AddChart2(gfStyle, gfType, gfLeft, gfTop, gfWidth, gfHeight).Chart
        ws.Shapes(ws.Shapes.Count).name = gfNome
        With gf
            .BarShape = gfBarShape
            .SeriesCollection.NewSeries
            ctSeries = gf.FullSeriesCollection.Count
            .FullSeriesCollection(ctSeries).name = gfNomeSerie 'renomeia a série (inicial)
            .ChartTitle.Text = gfTitulo
        End With
    Else 'Se o gráfico existe então apenas cria a série
        Set gf = ws.Shapes(gfNome).Chart
        With gf
            .SeriesCollection.NewSeries
            ctSeries = gf.FullSeriesCollection.Count
            .FullSeriesCollection(ctSeries).name = gfNomeSerie 'renomeia a série
        End With
    End If
    
    'Preenchendo dados do gráfico
    PreencherGrafico gfValores, gfRotulos, gf, gfNomeSerie
    
    'Formatações de gráfico
    With gf
        .SetElement msoElementDataLabelShow
        .ChartTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue
        '.Parent.RoundedCorners = True
        With .Walls.Format.Fill
            .Solid
            .ForeColor.RGB = RGB(255, 242, 204)
        End With
        With .Floor.Format.Fill
            .Solid
            .ForeColor.RGB = RGB(255, 242, 204)
        End With
        With .Axes(xlValue).MajorGridlines.Format.Line
            .ForeColor.RGB = RGB(197, 197, 197)
        End With
        .Parent.Placement = xlFreeFloating
    End With

End Sub
Sub GradientBKForm(usfX As Object, usfH As Double, usfW As Double, arCores As Variant)
    Dim tempChart As Chart
    Dim tempSheet As Worksheet
    Dim gradientBmp As Object
    Dim imgPath As String
    Dim imgx As MSForms.Image
    Dim i As Long
    Dim ct As Long
    
    OtimizarIniciar
    
    'Exclui plan temporária
    excluirSheet "wsTempShape"
    
    ' Cria um ChartObject temporário em uma planilha temporária
    Set tempSheet = ThisWorkbook.Worksheets.Add
    tempSheet.name = "wsTempShape"
    
    Set tempChart = tempSheet.ChartObjects.Add(left:=0, Width:=300, top:=0, Height:=200).Chart
    
    ' Configura o gráfico com um preenchimento degradê de duas cores
    With tempChart.Shapes.AddShape(msoShapeRectangle, 0, 0, tempChart.ChartArea.Width, tempChart.ChartArea.Height)
        .Fill.TwoColorGradient msoGradientHorizontal, UBound(arCores) - LBound(arCores) + 1
        For i = LBound(arCores) To UBound(arCores)
            ct = ct + 1
            .Fill.GradientStops(ct).Color.RGB = arCores(i)
        Next i
        .Line.Visible = msoFalse
    End With
    
    ' Ajusta o tamanho da forma para preencher completamente o gráfico
    With tempChart.Shapes(1)
        .Width = tempChart.ChartArea.Width * 2
        .Height = tempChart.ChartArea.Height * 2
    End With
    
    imgPath = Environ$("TEMP") & "\temp_image.jpg"
    tempChart.Export imgPath, "JPG"

    'Inserindo imagem no userform
    Set imgx = usfX.Controls.Add("Forms.Image.1", "ImageTemp", True)

    With imgx
        .Height = usfH
        .Width = usfW
        .Picture = LoadPicture(imgPath)
        .PictureSizeMode = fmPictureSizeModeStretch
    End With
    
    OrdemProfundidadeControle 0, usfX, Array("ImageTemp")
    
    ' Limpa os objetos temporários e exclui a imagem temporária
    Set gradientBmp = Nothing
    excluirSheet "wsTempShape"
    Kill imgPath ' Exclui a imagem temporária
    OtimizarFinalizar
End Sub
Sub GradientBKFrames(usfX As Object, arCores As Variant)
    Dim tempChart As Chart
    Dim tempSheet As Worksheet
    Dim gradientBmp As Object
    Dim imgPath As String
    Dim imgx As MSForms.Image
    Dim i As Long
    Dim ct As Long
    Dim c As Control
    
    
    'Formata o background color de todos os frames do formulário
    OtimizarIniciar
    
    'Exclui plan temporária
    excluirSheet "wsTempShape"
    
    ' Cria um ChartObject temporário em uma planilha temporária
    Set tempSheet = ThisWorkbook.Worksheets.Add
    tempSheet.name = "wsTempShape"
    
    Set tempChart = tempSheet.ChartObjects.Add(left:=0, Width:=300, top:=0, Height:=200).Chart
    
    ' Configura o gráfico com um preenchimento degradê de duas cores
    With tempChart.Shapes.AddShape(msoShapeRectangle, 0, 0, tempChart.ChartArea.Width, tempChart.ChartArea.Height)
        .Fill.TwoColorGradient msoGradientHorizontal, UBound(arCores) - LBound(arCores) + 1
        For i = LBound(arCores) To UBound(arCores)
            ct = ct + 1
            .Fill.GradientStops(ct).Color.RGB = arCores(i)
        Next i
        .Line.Visible = msoFalse
    End With
    
    ' Ajusta o tamanho da forma para preencher completamente o gráfico
    With tempChart.Shapes(1)
        .Width = tempChart.ChartArea.Width * 2
        .Height = tempChart.ChartArea.Height * 2
    End With
    
    imgPath = Environ$("TEMP") & "\temp_image.jpg"
    tempChart.Export imgPath, "JPG"
    
    For Each c In usfX.Controls
        If TypeName(c) = "Frame" Then
            With c
                .Picture = LoadPicture(imgPath)
                .PictureSizeMode = fmPictureSizeModeStretch
            End With
        End If
    Next c
    
    ' Limpa os objetos temporários e exclui a imagem temporária
    Set gradientBmp = Nothing
    excluirSheet "wsTempShape"
    Kill imgPath ' Exclui a imagem temporária
    OtimizarFinalizar
End Sub
Sub OrdemProfundidadeControle(zOrdem As Long, usfX As Object, arBtIgnore As Variant)
    Dim c As Control
    For Each c In usfX.Controls
        If Not InArray(arBtIgnore, c.name) Then
            c.ZOrder zOrdem
        End If
    Next c
End Sub
