VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} main 
   Caption         =   "UserForm1"
   ClientHeight    =   10560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17784
   OleObjectBlob   =   "main.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'Configura��es de frame
Const frmTop = 48
Const frmLeft = 4.8
Const frmHeight = 470.4

'Vari�vel de controle
Public proceed As Boolean
Public isEdit As Boolean

'Parceiros
Public idParceiro As Long


Private Sub btAddParceiro_Click()
    'Insert
    isEdit = False
    
    'Proximo ID
    idParceiro = TabelaProximoId(tParceiros)
    
    'Variavel de controle
    proceed = False
    
    'Formul�rio
    Parceiros.Show
    
    'Atualiza exibi��o de tabela
    AtualizarTabelaParceiros
    
    'Mensagem de confirmaa��o
    If proceed Then MsgBox "Parceiro adicionado com sucesso!", vbInformation
    
End Sub

Private Sub btEditarParceiro_Click()
    'Insert
    isEdit = True
    
    'Verifica sele��o
    If Not ListViewIsSelected(lvParceiros) Then Exit Sub
    
    'Id Selecionado
    idParceiro = ListViewSelectedID(lvParceiros)
    
    'Variavel de controle
    proceed = False
    
    'Formul�rio
    Parceiros.Show
    
    'Atualiza exibi��o de tabela
    AtualizarTabelaParceiros
    
    'Mensagem de confirmaa��o
    If proceed Then MsgBox "Parceiro editado com sucesso!", vbInformation
End Sub

Private Sub btParceiros_Click()
    Dim arVisiveis As Variant
    Dim arFixos As Variant
    
    'Configura��es de exibi��o do frame
    arVisiveis = Array(frmParceiros.name)
    arFixos = Array(frmMenu.name)
    ExibirFrames arVisiveis, arFixos, Me, frmTop, frmHeight, frmLeft
    
    'Atualizar lista
    AtualizarTabelaParceiros
End Sub

Private Sub btRemoverParceiro_Click()
    'Verifica se h� registros selecionados
    If Not ListViewIsSelected(lvParceiros) Then Exit Sub
    
    'Voce tem certeza ?
    If Not AreUSure("Voc� tem certeza que deseja excluir os itens selecionados ?") Then Exit Sub
    
    'Captura os ids selecionados
    Dim arIds As Variant
    arIds = ColunaLViewSelecionados(lvParceiros)
    
    'Update para inativo
    UpdateTable arIds, "Id", Array("Status='inativo'"), tParceiros, True, ThisWorkbookFullPath
    
    'Atualiza exibi��o de tabela
    AtualizarTabelaParceiros
    
    'Confirma��o
    MsgBox "Os registros selecionados foram exclu�dos com sucesso!", vbInformation
End Sub

Private Sub UserForm_Initialize()

    'Formata��o do formul�rio
    FormatarFormulario Me
    
    'Caption
    Me.Caption = "Sistema SolarSheet v1.0"
    
    'Inicia na tabela parceiros
    btParceiros_Click
End Sub
Private Sub AtualizarTabelaParceiros()
    Dim arDados As Variant
    Dim tipo As String
    
    'Filtro parceiro
    If optClientes.value Then tipo = "Cliente"
    If optColaboradores.value Then tipo = "Colaborador"
    If optFornecedores.value Then tipo = "Fornecedor"
    
    'Lista
    arDados = ListaDadosParceiros(tipo)
    
    'Preencher
    PreencherTabela lvParceiros, arDados
    
    'Desmarcar lv
    On Error Resume Next
    lvParceiros.SelectedItem.Selected = False
    
End Sub
