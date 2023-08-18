VERSION 5.00
Begin VB.Form FormLogin 
   Caption         =   "Login"
   ClientHeight    =   4320
   ClientLeft      =   10575
   ClientTop       =   5865
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   9090
   Begin VB.CommandButton BtnSair 
      Caption         =   "Sair"
      Height          =   660
      Left            =   6060
      TabIndex        =   6
      Top             =   2340
      Width           =   2040
   End
   Begin VB.CommandButton BtnEntrar 
      Caption         =   "Entrar"
      Height          =   660
      Left            =   6045
      TabIndex        =   5
      Top             =   1470
      Width           =   2040
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
      Left            =   3735
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2325
      Width           =   1500
   End
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
      Left            =   3735
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1485
      Width           =   1500
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
      Height          =   660
      Left            =   975
      TabIndex        =   7
      Top             =   3435
      Width           =   7080
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
      Left            =   1740
      TabIndex        =   2
      Top             =   2355
      Width           =   1500
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
      Left            =   1770
      TabIndex        =   1
      Top             =   1530
      Width           =   1470
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Sistema do Leo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1560
      TabIndex        =   0
      Top             =   405
      Width           =   5820
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ADMIN As Integer ' 1 -> ADMIN, 2 -> NORMAL



Private Sub Form_Load()
    
    ' SE NÃO FOR ADMIN COMEÇA COMO 2
    ADMIN = 2
    
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

'Verificar tecla clicada no login
Private Sub TxtLogin_KeyPress(KeyAscii As Integer)
    
    On Error GoTo TrataErro
    
    If KeyAscii = 13 Then ' The ENTER key.
        
        'Verifica se o operador foi preenchido
        If TxtLogin.Text = "" Then
            MsgBox "Informe um operador"
        ElseIf TxtLogin.Text = 0 Then
            LblNomeOperador.Caption = "ADMIN"
            SendKeys ("{TAB}") ' Set the focus to the next control.
            KeyAscii = 0       ' Ignore this key.
        Else
            SendKeys "{tab}"   ' Set the focus to the next control.
            KeyAscii = 0       ' Ignore this key.
            
            'Debug para verificar o retorno da variável
            'Debug.Print "Return: " & cmd.Parameters("RetornoOperacao").Value
            'Debug.Print "Output: " & cmd.Parameters("OUTPUT").Value
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

    Dim SQL As String
    Dim validacao As Integer
    Dim rs As New ADODB.Recordset

    ' CHAMA A FUNÇÃO QUE VERIFICA SE O OPERADOR EXISTE
    If TxtLogin = "" Then
        
    Else
        validacao = verificaOperador(TxtLogin.Text)

        'Verifica o retorno da função (0 - existe; 1 - Inativo; 2 - Não cadastrado)
        'Retorno = 0
        If validacao = 0 Then
            'SELECT para puxar o nome do operador
            SQL = "SELECT nomeOperador FROM Operador WHERE IdOperador = " & TxtLogin.Text
            
            'Abre o recordset
            rs.Open SQL, cn, adOpenStatic
                'Adiciona o nome do operador no Label
                If rs.EOF = True Then
                
                Else
                    LblNomeOperador.Caption = rs("nomeOperador")
                End If
            rs.Close
        'Retorno = 1
        ElseIf validacao = 1 Then
            MsgBox "Operador Inativo"
            TxtLogin.SetFocus
            TxtLogin.Text = ""
            LblNomeOperador.Caption = ""
        'Retorno = 0
        Else
            MsgBox "Operador não Cadastrado"
            TxtLogin.SetFocus
            TxtLogin.Text = ""
            LblNomeOperador.Caption = ""
        End If
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

Private Sub BtnEntrar_Click()
    
    Dim CMD As New ADODB.Command
    Dim Parametros As New ADODB.parameter
    Dim rs As New ADODB.Recordset
    
    Dim SQL As String
    Dim nomeParametros As String
    Dim retorno As Integer
    
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
        retorno = CMD.Parameters("RetornoOperacao").Value
        
        If retorno = 0 Then
            
            If TxtLogin.Text = 0 Then
                ADMIN = 1 ' ADMINISTRADO
            Else
                SQL = "SELECT adminOperador FROM Operador WHERE IdOperador = " & TxtLogin.Text
                
                rs.Open SQL, cn, adOpenStatic
                    If rs("adminOperador") = True Then
                        ADMIN = 1 ' ADMINISTRADO
                    End If
                rs.Close
            End If
            MDIFormInicio.ID = TxtLogin.Text
            Unload Me
            MDIFormInicio.Show
        ElseIf retorno = 1 Then
            MsgBox "Senha Incorreta"
            TxtLogin.Text = ""
            TxtSenha.Text = ""
            LblNomeOperador.Caption = ""
            TxtLogin.SetFocus
        ElseIf retorno = 2 Then
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
    End
End Sub
