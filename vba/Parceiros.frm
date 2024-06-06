VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Parceiros 
   Caption         =   "UserForm1"
   ClientHeight    =   7956
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4500
   OleObjectBlob   =   "Parceiros.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Parceiros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'Id do parceiro
Public thisID As Long

Private Sub btCale_Click()
    ShowCalPt Data_Cadasto, "Data de Cadastro"
End Sub

Private Sub btCancelar_Click()
    main.proceed = False
    Unload Me
End Sub

Private Sub btSalvar_Click()
    Dim arOb As Variant
    Dim arColValues As Variant
    Dim arValues As Variant, arColumns As Variant
    
    'Campos obrigatórios
    arOb = Array(Nome, Uf, Cidade, Tel_Contato, Tipo_Parceiro)
    
    'Valida formulario
    If Not ValidarFormulario(Me, arOb) Then Exit Sub
    
    'Verifica se o tipo é "Colaborador" e se a Função foi preenchida
    If Tipo_Parceiro = "Colaborador" And Cargo_Colaborador.value = "" Then
        MsgBox "O campo 'Cargo/Função' é obrigatório para colaboradores!", vbExclamation
        Exit Sub
    End If
    
    'Salva os dados
    If main.isEdit Then
        arColValues = ColunaValorUpdate(Me, tParceiros)
        UpdateTable Array(thisID), "Id", arColValues, tParceiros, True, ThisWorkbookFullPath
    Else
        ColunaValorInsertInto Me, tParceiros, arColumns, arValues, "Id", "Status", thisID
        InsertIntoTable arColumns, arValues, tParceiros, True, ThisWorkbookFullPath
    End If
    
    'Libera proceed
    main.proceed = True
    
    'Fecha o formulário
    Unload Me
End Sub

Private Sub Cargo_Colaborador_Change()
    PreencherTextBoxComID Cargo_Colaborador, Id_Funcao, 0
End Sub

Private Sub Tipo_Parceiro_Change()
    Dim tipo As String
    
    tipo = Tipo_Parceiro.value
    If tipo <> "Colaborador" Then
        Cargo_Colaborador.Visible = False
        Label12.Visible = False
    Else
        Cargo_Colaborador.Visible = True
        Label12.Visible = True
    End If
End Sub

Private Sub Uf_Change()
    Dim selectedUF As String
    
    selectedUF = Uf.value
    
    Cidade.value = ""
    Cidade.List = ListaCidades(selectedUF)
End Sub

Private Sub UserForm_Initialize()
    
    'Comboboxes
    Uf.List = ListaUF()
    Cargo_Colaborador.List = ListaFuncoesID()
    Tipo_Parceiro.List = Array("Cliente", "Fornecedor", "Colaborador")
    
    'CurID
    thisID = main.idParceiro
    
    'Preencher formulario com os dados existentes do id
    PreencherDadosFormularioSELECT Me, thisID, "Id", tParceiros
    
    'Iniciais
    If (Data_Cadastro.value = "" Or Data_Cadastro = "00:00:00") Then Data_Cadastro = Date
  
    'Formatando formulario
    FormatarFormulario Me
    
    'Caption
    Me.Caption = "Parceiros"
    
End Sub
