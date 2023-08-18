Attribute VB_Name = "SPsGlobais"
Public Function VerificaEstoqueProduto(IDProduto As Integer)

    Dim CMD As New ADODB.Command
    Dim Parametros As New ADODB.Parameter
    
    Dim retorno As Long
    
    With CMD
    Set .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "VerificaEstoqueProduto"
    End With
    
    'Parametro 1
    nomeParametros = "RetornoOperacao"
    Set Parametros = CMD.CreateParameter(nomeParametros, adInteger, adParamReturnValue) 'Cria um parametro utilizando o nome, tipo e metodo(RETORNO/ENVIO/OUTPUT)
        CMD.Parameters.Append Parametros '"Adiciona" o parametro dentro do Command
        CMD.Parameters(nomeParametros).Value = -1 'Seta o valor do parametro (valor aleatório para teste)
    'Parametro 2
    nomeParametros = "OUTPUT"
    Set Parametros = CMD.CreateParameter(nomeParametros, adInteger, adParamOutput) 'Cria um parametro utilizando o nome, tipo e metodo(RETORNO/ENVIO/OUTPUT)
        CMD.Parameters.Append Parametros '"Adiciona" o parametro dentro do Command
        CMD.Parameters(nomeParametros).Value = -1 'Seta o valor do parametro (valor aleatório para teste)
    'Parametro 3
    nomeParametros = "IdProduto"
    Set Parametros = CMD.CreateParameter(nomeParametros, adInteger, adParamInput) 'Cria um parametro utilizando o nome, tipo e metodo(RETORNO/ENVIO/OUTPUT)
        CMD.Parameters.Append Parametros '"Adiciona" o parametro dentro do Command
        CMD.Parameters(nomeParametros).Value = IDProduto 'Seta o valor do parametro
        
    'Executa o Command (SP)
    CMD.Execute
    'Adiciona o retorno da SP na variavel retorno
    retorno = CMD.Parameters("RetornoOperacao").Value

    ' DEVOLVE A QUANTIDADE DO ESTOQUE PARA A TELA QUE CHAMOU
    VerificaEstoqueProduto = retorno

End Function

Public Function verificaOperador(ID As Integer)

    Dim rs As New ADODB.Recordset
    Dim CMD As New ADODB.Command
    Dim validacao As Integer
    
    'Inicia as informações no Command
    CMD.ActiveConnection = cn 'Conexão
    CMD.CommandType = adCmdStoredProc 'Tipo de procedimento (SP)
    CMD.CommandText = "VerificaOperador" 'Nome da SP
    
    ' PARAMETROS PARA A SP
    CMD.Parameters.Append CMD.CreateParameter("RetornoOperacao", adSmallInt, adParamReturnValue, , 99)
    CMD.Parameters.Append CMD.CreateParameter("OUTPUT", adSmallInt, adParamOutput, , 99)
    CMD.Parameters.Append CMD.CreateParameter("ID", adInteger, adParamInput, , ID)
    
    ' EXECUTA A SP
    CMD.Execute

    ' PEGA O RETORNO DA SP
    validacao = CMD.Parameters("RetornoOperacao").Value

    ' DEVOLVE O RETORNO PARA QUEM CHAMOU
    verificaOperador = validacao

End Function
