VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public WithEvents tBox          As MSForms.TextBox
Attribute tBox.VB_VarHelpID = -1
Public WithEvents cBox          As MSForms.ComboBox
Attribute cBox.VB_VarHelpID = -1
Public WithEvents lbl           As MSForms.Label
Attribute lbl.VB_VarHelpID = -1
Public WithEvents optButton     As MSForms.OptionButton
Attribute optButton.VB_VarHelpID = -1
Public WithEvents chkButton     As MSForms.CheckBox
Attribute chkButton.VB_VarHelpID = -1
Public WithEvents frm           As MSForms.Frame
Attribute frm.VB_VarHelpID = -1
Public WithEvents cmdButton     As MSForms.CommandButton
Attribute cmdButton.VB_VarHelpID = -1


Public arControlsSetup          As Variant  'Array (0 to n, 0 to 2) |Nome Controle, Tipo, Label Caption|
Public ParentForm               As MSForms.UserForm


Public Sub BuildForm()
    Dim i As Long, j As Long
    Dim typ As String, name As String
    Dim cTop As Double, cLeft As Double
    Dim topIni As Double, LeftIni As Double
    
    topIni = 4.8
    LeftIni = 4.8
    
    For i = LBound(arControlsSetup) To UBound(arControlsSetup)
        'Nome e tipo do botão
        typ = DicControlsSetup.Item(ky)
        typ = "Forms." & typ & ".1"
        name = ky
        
        'Configurando posição no formulário
        
        
        'Cria os botões
        
    Next i
        
End Sub

