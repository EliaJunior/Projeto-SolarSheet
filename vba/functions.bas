Attribute VB_Name = "functions"
Option Explicit
Option Base 0
Function ListaUF() As Variant
    Dim stSQL As String
    Dim arOut As Variant
    Dim dbFolder As String
    Dim fileName As String
    
    dbFolder = ThisWorkbookPath & "\db"
    fileName = "cidades_brasil.csv"

    stSQL = "SELECT " & _
                "DISTINCT([UF]) " & _
            "FROM [" & fileName & "] " & _
            "ORDER BY " & _
                "[UF] ASC"
            
    
    arOut = ConsultaSQL(stSQL, False, True, dbFolder, True)

    If isArrayNotEmpty(arOut) Then ListaUF = arOut Else ListaUF = arEscape()
End Function
Function ListaCidades(Uf As String) As Variant
    Dim stSQL As String
    Dim arOut As Variant
    Dim dbFolder As String
    Dim fileName As String
    
    dbFolder = ThisWorkbookPath & "\db"
    fileName = "cidades_brasil.csv"

    stSQL = "SELECT " & _
                "[Municipio] " & _
            "FROM [" & fileName & "] " & _
            "WHERE " & _
                "[UF] = '" & Uf & "' " & _
            "ORDER BY " & _
                "[Municipio] ASC"
            
    
    arOut = ConsultaSQL(stSQL, False, True, dbFolder, True)

    If isArrayNotEmpty(arOut) Then ListaCidades = arOut Else ListaCidades = arEscape()
End Function
Function ListaFuncoes() As Variant
    Dim stSQL As String
    Dim arOut As Variant
    Dim tb As String
    
    tb = TabelaRefSQL(tFuncoes, 2, "t")

    stSQL = "SELECT " & _
                "[Funcao] " & _
            "FROM " & tb & _
            "WHERE " & _
                "[Status] = 'ativo' " & _
            "ORDER BY " & _
                "[Funcao] ASC"
   
    arOut = ConsultaSQL(stSQL, False, True, ThisWorkbookFullPath, False)

    If isArrayNotEmpty(arOut) Then ListaFuncoes = arOut Else ListaFuncoes = arEscape()
End Function
Function arEscape() As Variant
    Dim ar(0 To 0, 0 To 0)
    
    ar(0, 0) = "Sem registros"
    arEscape = ar
End Function
