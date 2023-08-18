VERSION 5.00
Begin VB.Form FormSolicitaAcesso 
   Caption         =   "Solicitação de Acesso"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtLogin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2685
      MaxLength       =   4
      TabIndex        =   0
      Top             =   315
      Width           =   1500
   End
   Begin VB.TextBox TxtSenha 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      IMEMode         =   3  'DISABLE
      Left            =   2685
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1155
      Width           =   1500
   End
   Begin VB.CommandButton BtnEntrar 
      Caption         =   "Entrar"
      Height          =   660
      Left            =   5025
      TabIndex        =   2
      Top             =   240
      Width           =   2040
   End
   Begin VB.CommandButton BtnSair 
      Caption         =   "Sair"
      Height          =   660
      Left            =   5025
      TabIndex        =   4
      Top             =   1005
      Width           =   2040
   End
   Begin VB.Label LblNomeOperador 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   630
      Left            =   630
      TabIndex        =   6
      Top             =   2160
      Width           =   6150
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Login:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   720
      TabIndex        =   5
      Top             =   360
      Width           =   1470
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   690
      TabIndex        =   3
      Top             =   1185
      Width           =   1500
   End
End
Attribute VB_Name = "FormSolicitaAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ADMIN As Integer
Public FORMULARIO As String

Private Sub BtnEntrar_Click()
    
    Dim CMD As New ADODB.Command
    Dim Parametros As New ADODB.parameter
    Dim rs As New ADODB.Recordset
    
    Dim SQL As String
    Dim nomeParametros As String
    Dim Retorno As Integer
    
    On Error GoTo TrataErro
        
    'Verifica se os campos estão preenchidos
    If TxtLogin.Text = "" Then
        MsgBox "Campo Operador não pode estar vazio"
        TxtLogin.SetFocus
    ElseIf TxtSenha.Text = "" Then
        MsgBox "Campo Senha não pode estar vazio"
        TxtSenha.SetFocus
    Else
        With CMD
            Set .ActiveConnection = cn
            .CommandType = adCmdStoredProc
            .CommandText = "VerificarLogin"
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
        nomeParametros = "ID"
        Set Parametros = CMD.CreateParameter(nomeParametros, adInteger, adParamInput) 'Cria um parametro utilizando o nome, tipo e metodo(RETORNO/ENVIO/OUTPUT)
            CMD.Parameters.Append Parametros '"Adiciona" o parametro dentro do Command
            CMD.Parameters(nomeParametros).Value = TxtLogin.Text 'Seta o valor do parametro
        'Parametro 4
        nomeParametros = "SENHA"
        Set Parametros = CMD.CreateParameter(nomeParametros, adInteger, adParamInput) 'Cria um parametro utilizando o nome, tipo e metodo(RETORNO/ENVIO/OUTPUT)
            CMD.Parameters.Append Parametros '"Adiciona" o parametro dentro do Command
            CMD.Parameters(nomeParametros).Value = TxtSenha.Text 'Seta o valor do parametro
        'Executa o Command (SP)
        CMD.Execute
        'Adiciona o retorno da SP na variavel retorno
        Retorno = CMD.Parameters("RetornoOperacao").Value
        
        If Retorno = 0 Then
            
            SQL = "SELECT adminOperador FROM Operador WHERE IdOperador = " & TxtLogin.Text
            
            rs.Open SQL, cn, adOpenStatic
                If rs.EOF = True And TxtLogin.Text = 0 Then
                    ADMIN = 1
                    MsgBox ("Liberação efetuada com sucesso!")
                    ' CHAMA A FUNÇÃO QUE VERIFICA QUAL TELA DEVE FAZER A LIBERAÇÃO
                    verificarFormulario
                Else
                    If rs("adminOperador") = 1 Then
                        ADMIN = 1
                        MsgBox ("Liberação efetuada com sucesso!")
                        ' CHAMA A FUNÇÃO QUE VERIFICA QUAL TELA DEVE FAZER A LIBERAÇÃO
                        verificarFormulario
                    Else
                        MsgBox ("Operador " & TxtLogin.Text & " não tem permissão para efetuar a liberação")
                        TxtLogin.Text = ""
                        TxtSenha.Text = ""
                        LblNomeOperador.Caption = ""
                        TxtLogin.SetFocus
                    End If
                End If
            rs.Close
        ElseIf Retorno = 1 Then
            MsgBox "Senha Incorreta"
            TxtLogin.Text = ""
            TxtSenha.Text = ""
            LblNomeOperador.Caption = ""
            TxtLogin.SetFocus
        ElseIf Retorno = 2 Then
            MsgBox "Operador " & TxtLogin.Text & " inativo"
            TxtLogin.Text = ""
            TxtSenha.Text = ""
            TxtLogin.SetFocus
        Else
            MsgBox "Operador " & TxtLogin.Text & " não cadastrado"
            TxtLogin.Text = ""
            TxtSenha.Text = ""
            TxtLogin.SetFocus
        End If
    End If
Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro ao tentar logar: " & TxtLogin & vbCrLf & _
           "Erro número : " & Err.Number & vbCrLf & _
           "Detalhes : " & Err.Description
End Sub

Private Sub BtnSair_Click()
    
    finalizaForm
    
End Sub

Private Sub Form_Load()
    
    ' NÃO ADMIN
    ADMIN = 2
    
    ' ALTERA O NOME DO FORM
    FormSolicitaAcesso.Caption = "Solicitação de Acesso - " & FORMULARIO
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' ATIVA NOVAMENTE O MDI
    MDIFormInicio.Enabled = True
End Sub

Private Sub TxtLogin_Change()
    
    TxtSenha.Text = ""
    
End Sub

Private Sub TxtLogin_GotFocus()

    With TxtLogin
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub TxtLogin_KeyPress(KeyAscii As Integer)
    
    On Error GoTo TrataErro
    
    If KeyAscii = 13 Then ' The ENTER key.
        'Verifica se o operador foi preenchido
        If TxtLogin.Text = "" Then
            MsgBox "Informe um operador"
        ElseIf TxtLogin.Text = 0 Then
            LblNomeOperador.Caption = "Administrador"
            SendKeys ("{TAB}")
            KeyAscii = 0
        Else
            SendKeys ("{TAB}")
            KeyAscii = 0
        End If
    End If
    
    'ele só vai aceitar números e o backspace.
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
       KeyAscii = 0
    End If
    
Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro ao localizar o operador: " & TxtLogin & vbCrLf & _
           "Erro número : " & Err.Number & vbCrLf & _
           "Detalhes : " & Err.Description
End Sub

Private Sub TxtLogin_LostFocus()

    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim validacao As Integer

    If TxtLogin.Text = "" Then
        TxtLogin.Text = 0
        validacao = 0
    Else
        validacao = verificaOperador(TxtLogin.Text)
    End If
            
    'Verifica o retorno da função (0 - existe; 1 - Inativo; 2 - Não cadastrado)
    'Retorno = 0
    If validacao = 0 Then
        'MsgBox "Operador existe"
        'SELECT para puxar o nome do operador
        SQL = "SELECT nomeOperador FROM Operador WHERE IdOperador = " & TxtLogin.Text
            
        'Abre o recordset
        rs.Open SQL, cn, adOpenStatic
            If rs.EOF = True Then
                
            Else
                'Adiciona o nome do operador no Label
                LblNomeOperador.Caption = rs("nomeOperador")
            End If
        rs.Close
    'Retorno = 1
    ElseIf validacao = 1 Then
        MsgBox "Operador Inativo"
        TxtLogin.Text = ""
        TxtLogin.SetFocus
        LblNomeOperador.Caption = ""
    'Retorno = 0
    Else
        MsgBox "Operador não Cadastrado"
        TxtLogin.Text = ""
        TxtLogin.SetFocus
        LblNomeOperador.Caption = ""
    End If

End Sub

Private Sub TxtSenha_GotFocus()

    With TxtSenha
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtSenha_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then ' The ENTER key.
        SendKeys "{tab}"
        KeyAscii = 0
    End If
    
    'ele só vai aceitar números e o backspace.
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
       KeyAscii = 0
    End If
End Sub

Private Function finalizaForm()

    MDIFormInicio.Enabled = True
    Unload Me

End Function

Private Function verificarFormulario()
    
    ' LIBERAÇÃO PEDIDOS
    If FORMULARIO = "FormPedidos" Then
        FormPedidos.TxtValor.Enabled = True
        FormPedidos.TxtValorItem.Enabled = True
    ' LIBERAÇÃO DE ENTRADA
    ElseIf FORMULARIO = "FormEntrada" Then
        FormEntrada.TxtCusto.Enabled = True
        FormEntrada.TxtCustoItem.Enabled = True
    ' LIBERAÇÃO CONFIGURAÇÕES
    ElseIf FORMULARIO = "MDIFormInicio" Then
        FormConfiguracoesGerais.Show
    End If

    finalizaForm
    
End Function
