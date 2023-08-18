VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FormCadastroOperador 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro Operador"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtNome 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1770
      TabIndex        =   2
      Top             =   1185
      Width           =   9990
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   6240
      TabIndex        =   12
      Top             =   60
      Width           =   5670
      Begin VB.CommandButton BtnAlterar 
         Caption         =   "Alterar"
         Height          =   660
         Left            =   4395
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   225
         Width           =   1000
      End
      Begin VB.CommandButton BtnConfirmar 
         Caption         =   "Confirmar"
         Height          =   660
         Left            =   345
         TabIndex        =   15
         Top             =   225
         Width           =   1000
      End
      Begin VB.CommandButton BtnCancelar 
         Caption         =   "Cancelar"
         Height          =   660
         Left            =   1710
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   225
         Width           =   1000
      End
      Begin VB.CommandButton BtnNovo 
         Caption         =   "Novo"
         Height          =   660
         Left            =   3045
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   225
         Width           =   1000
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4620
      Left            =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1905
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   8149
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Informações"
      TabPicture(0)   =   "FormCadastroOperador.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ChkInativo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TxtSenha"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ChkVisualizarSenha"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ChkADMIN"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Liberações (Telas)"
      TabPicture(1)   =   "FormCadastroOperador.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ChkFormCadastroOperador"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ChkFormCadastroProduto"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "ChkFormTabelaUsuario"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ChkFormPedidos"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "ChkFormCadastroCategoria"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ChkFormCadastroCliente"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "ChkFormEntrada"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "ChkFormCadastroFormaPGTO"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin VB.CheckBox ChkFormCadastroFormaPGTO 
         Caption         =   "Cadastro Forma PGTO"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -74700
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2400
         Width           =   3375
      End
      Begin VB.CheckBox ChkFormEntrada 
         Caption         =   "Formulario Entrada"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -74745
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3915
         Width           =   2805
      End
      Begin VB.CheckBox ChkFormCadastroCliente 
         Caption         =   "Cadastro Cliente"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -74700
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1425
         Width           =   2910
      End
      Begin VB.CheckBox ChkFormCadastroCategoria 
         Caption         =   "Cadastro Categoria"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -74700
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   915
         Width           =   2880
      End
      Begin VB.CheckBox ChkFormPedidos 
         Caption         =   "Formulario Pedido"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -74700
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3405
         Width           =   2700
      End
      Begin VB.CheckBox ChkFormTabelaUsuario 
         Caption         =   "Listagem Clientes"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -74700
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2895
         Width           =   2655
      End
      Begin VB.CheckBox ChkFormCadastroProduto 
         Caption         =   "Cadastro Produto"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -74700
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1905
         Width           =   2910
      End
      Begin VB.CheckBox ChkFormCadastroOperador 
         Caption         =   "Cadastro Operador"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -74700
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   405
         Width           =   2850
      End
      Begin VB.CheckBox ChkADMIN 
         Caption         =   "Administrador"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   480
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1275
         Width           =   2250
      End
      Begin VB.CheckBox ChkVisualizarSenha 
         Caption         =   "Visualizar senha"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4200
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   555
         Width           =   2565
      End
      Begin VB.TextBox TxtSenha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         IMEMode         =   3  'DISABLE
         Left            =   1710
         MaxLength       =   4
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   525
         Width           =   2280
      End
      Begin VB.CheckBox ChkInativo 
         Caption         =   "Inativo"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   480
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1530
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Senha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   360
         TabIndex        =   9
         Top             =   570
         Width           =   1200
      End
   End
   Begin VB.CommandButton BtnAvancaRegistro 
      Caption         =   ">"
      Height          =   555
      Left            =   5190
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   345
      Width           =   675
   End
   Begin VB.CommandButton BtnVoltaRegistro 
      Caption         =   "<"
      Height          =   555
      Left            =   4260
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   345
      Width           =   675
   End
   Begin VB.CommandButton BtnPesquisarOperador 
      Caption         =   "->"
      Height          =   450
      Left            =   1785
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   390
      Width           =   690
   End
   Begin VB.TextBox TxtIdOperador 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2745
      TabIndex        =   1
      Top             =   330
      Width           =   1260
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   390
      TabIndex        =   23
      Top             =   1230
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   285
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "FormCadastroOperador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public STATUS As Integer  ' 0 -> CADASTRO NOVO / 1 -> CADASTRO JÁ FINALIZADO / 2 - EDIÇÃO DE CADASTRO JÁ FINALIZADO

Private Sub BtnAlterar_Click()

    ' ALTERA PARA O STATUS 2 -> ALTERAÇÃO DE CADASTRO JÁ FINALIZADO
    STATUS = 2
    ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
    preencheOperador (TxtIdOperador.Text)

End Sub

Private Sub BtnAvancaRegistro_Click()

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim ID As Integer
    
    SQL = "DECLARE @ID INT " & _
          "SET @ID = " & TxtIdOperador.Text & " " & _
          "SELECT TOP 1 IdOperador FROM Operador WHERE IdOperador > @ID ORDER BY IdOperador"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            ID = rs("IdOperador")
        Else
            ID = TxtIdOperador.Text
        End If
    rs.Close
    
    TxtIdOperador.Text = ID
    preencheOperador (ID)

End Sub

Private Sub BtnAvancaRegistro_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnCancelar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnPesquisarOperador_Click()
    FormBuscaOperador.FORMULARIO = "FormCadastroOperador"
    FormBuscaOperador.Show
    FormBuscaOperador.SetFocus
    MDIFormInicio.Enabled = False
End Sub

Private Sub BtnPesquisarOperador_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then ' The ENTER key.
        SendKeys ("{TAB}")
    ElseIf KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnConfirmar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnNovo_Click()
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim ID As Integer
    
    SQL = "SELECT MAX(IdOperador)+1 AS Operador FROM Operador"
    
    rs.Open SQL, cn, adOpenStatic
        ID = rs("Operador")
        TxtIdOperador.Text = rs("Operador")
    rs.Close
    
    ' SETO O STATUS DA TELA PARA 0 -> CADASTRO NOVO
    STATUS = 0
    preencheOperador (ID)
    
    ' ******* COLOCA O FOCO NA PRIMEIRA ABA DO SSTab *********
    SSTab1.Tab = 0
End Sub

Private Sub BtnNovo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnVoltaRegistro_Click()

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim ID As Integer
    
    SQL = "DECLARE @ID INT " & _
          "SET @ID = " & TxtIdOperador.Text & " " & _
          "SELECT TOP 1 IdOperador FROM Operador WHERE IdOperador < @ID ORDER BY IdOperador DESC"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            ID = rs("IdOperador")
        Else
            ID = TxtIdOperador.Text
        End If
    rs.Close
    
    TxtIdOperador.Text = ID
    preencheOperador (ID)

End Sub

Private Sub BtnVoltaRegistro_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub ChkInativo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub ChkVisualizarSenha_Click()

    If ChkVisualizarSenha.Value = Checked Then
        TxtSenha.PasswordChar = ""
    Else
        TxtSenha.PasswordChar = "*"
    End If

End Sub

Private Sub Form_Load()
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim ID As Integer
    
    SQL = "DECLARE @maxID INT " & _
          "SELECT @maxID = MAX(IdOperador) FROM Operador " & _
          "SELECT @maxID AS Operador"
    
    rs.Open SQL, cn, adOpenStatic
        If IsNull(rs("Operador")) Then
            ' DEFINE O STATUS PARA 1 -> CLIENTE NOVO
            STATUS = 0
            ID = 1
            TxtIdOperador.Text = 1
        Else
            ID = rs("Operador")
            ' DEFINE O STATUS PARA 1 -> CLIENTE JÁ CADASTRADO
            STATUS = 1
        End If
    rs.Close
    
    ' ******* COLOCA O FOCO NA PRIMEIRA ABA DO SSTab *********
    SSTab1.Tab = 0
    
    preencheOperador (ID)
    
End Sub

Private Sub BtnCancelar_Click()
    
    finalizaForm
    
End Sub

Private Sub BtnConfirmar_Click()

    'On Error GoTo TrataErro

    Dim SQL As String
    Dim StatusOperador As Integer
    Dim Administrador As Integer
    Dim validacao As Integer
    Dim CMD As New ADODB.Command
    Dim rs As New ADODB.Recordset
    Dim telaOperador, telaCategoria, telaCliente, telaProduto, telaFormaPGTO, telaTabelaUsuario, telaPedidos, telaEntrada As Boolean
    Dim ID As Integer
    Dim Operador As Integer
    
    If TxtNome.Text = "" Then
        MsgBox "Preencha o campo Nome!"
        TxtNome.SetFocus
    ElseIf TxtSenha.Text = "" Then
        MsgBox ("Preencha o campo Senha!")
        TxtSenha.SetFocus
    Else
        ' SE ESTIVER TUDO COMPLETO EXECUTA
        
        ' VERIFICA O STATUS PARA PASSAR O ID CORRETO
        If STATUS = 0 Then
            Operador = -1
        ElseIf STATUS = 2 Then
            Operador = TxtIdOperador.Text
        End If
        
        ' VERIFICA O INATIVO
        If ChkInativo.Value = Checked Then
            StatusOperador = True
        Else
            StatusOperador = False
        End If
        
        ' VERIFICA O ADMINISTRADOR
        If ChkADMIN.Value = Checked Then
            Administrador = True
        Else
            Administrador = False
        End If
        
        ' VERIFICA A TELA CADASTRO DE OPERADOR
        If ChkFormCadastroOperador.Value = Checked Then
            telaOperador = True
        Else
            telaOperador = False
        End If
        
        ' VERIFICA A TELA CADASTRO DE CATEGORIA
        If ChkFormCadastroCategoria.Value = Checked Then
            telaCategoria = True
        Else
            telaCategoria = False
        End If
        
        ' VERIFICA A TELA CADASTRO DE CLIENTE
        If ChkFormCadastroCliente.Value = Checked Then
            telaCliente = True
        Else
            telaCliente = False
        End If
        
        ' VERIFICA A TELA CADASTRO PRODUTO
        If ChkFormCadastroProduto.Value = Checked Then
            telaProduto = True
        Else
            telaProduto = False
        End If
        
        ' VERIFICA A TELA CADASTRO FORMA PGTO
        If ChkFormCadastroFormaPGTO.Value = Checked Then
            telaFormaPGTO = True
        Else
            telaFormaPGTO = False
        End If
        
        ' VERIFICA A TELA TABELA DE CLIENTES
        If ChkFormTabelaUsuario.Value = Checked Then
            telaTabelaUsuario = True
        Else
            telaTabelaUsuario = False
        End If
        
        ' VERIFICA A TELA DE PEDIDOS
        If ChkFormPedidos.Value = Checked Then
            telaPedidos = True
        Else
            telaPedidos = False
        End If
        
        If ChkFormEntrada.Value = Checked Then
            telaEntrada = True
        Else
            telaEntrada = False
        End If
    
        ' PASSA A CONEXÃO PARA O COMMAND
        CMD.ActiveConnection = cn
        CMD.CommandText = "cadastroOperador"
        CMD.CommandType = adCmdStoredProc
        
        ' PASSA OS PARAMETROS PARA O COMMAND
        CMD.Parameters.Append CMD.CreateParameter("RetornoOperacao", adInteger, adParamReturnValue, , 99)
        CMD.Parameters.Append CMD.CreateParameter("OUTPUT", adInteger, adParamOutput, , 99)
        CMD.Parameters.Append CMD.CreateParameter("ID", adInteger, adParamInput, , Operador)
        CMD.Parameters.Append CMD.CreateParameter("Nome", adVarChar, adParamInput, 50, TxtNome.Text)
        CMD.Parameters.Append CMD.CreateParameter("Senha", adInteger, adParamInput, 14, TxtSenha.Text)
        CMD.Parameters.Append CMD.CreateParameter("Admin", adBoolean, adParamInput, 50, Administrador)
        CMD.Parameters.Append CMD.CreateParameter("Status", adBoolean, adParamInput, 1, StatusOperador)
        CMD.Parameters.Append CMD.CreateParameter("TelaCadOperador", adBoolean, adParamInput, 1, telaOperador)
        CMD.Parameters.Append CMD.CreateParameter("TelaCadCategoria", adBoolean, adParamInput, 1, telaCategoria)
        CMD.Parameters.Append CMD.CreateParameter("TelaCadCliente", adBoolean, adParamInput, 1, telaCliente)
        CMD.Parameters.Append CMD.CreateParameter("TelaCadProduto", adBoolean, adParamInput, 1, telaProduto)
        CMD.Parameters.Append CMD.CreateParameter("TelaCadFormaPGTO", adBoolean, adParamInput, 1, telaFormaPGTO)
        CMD.Parameters.Append CMD.CreateParameter("TelaTabelaUsuario", adBoolean, adParamInput, 1, telaTabelaUsuario)
        CMD.Parameters.Append CMD.CreateParameter("TelaPedidos", adBoolean, adParamInput, 1, telaPedidos)
        CMD.Parameters.Append CMD.CreateParameter("TelaEntrada", adBoolean, adParamInput, 1, telaEntrada)
        
        CMD.Execute
        
        validacao = CMD.Parameters("RetornoOperacao").Value
        
        If STATUS = 0 Then
            If validacao = 1 Then
                MsgBox ("Cliente cadastrado com sucesso!")
                
                ' FINALIZOU COM SUCESSO, PEGA O ID
                'ID = TxtIdOperador.Text
                        
                ' DEFINE O STATUS PARA 1 -> REGISTRO JÁ CADASTRADO
                'STATUS = 1
        
                'preencheOperador (ID)
            ElseIf validacao = 2 Then
                MsgBox ("Ocorreu algum erro ao tentar cadastrar o Operador!")
            End If
        ElseIf STATUS = 2 Then
            If validacao = 0 Then
                ' ALTEROU O CLIENTE COM SUCESSO
                MsgBox ("Cliente alterado com sucesso!")
            ElseIf validacao = 1 Then
                MsgBox ("Ocorreu algum erro ao tentar alterar o Operador!")
                Exit Sub
            End If
        End If
        
        ' FINALIZOU COM SUCESSO, PEGA O ID
        ID = TxtIdOperador.Text
                
        ' DEFINE O STATUS PARA 1 -> REGISTRO JÁ CADASTRADO
        STATUS = 1
        
        preencheOperador (ID)
    End If
Exit Sub
TrataErro:
    MsgBox "Algum erro ocorreu ao tentar finalizar o registro - " & TxtIdOperador.Text & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' CHAMA A FUNÇÃO PARA ATUALIZAR AS PERMISSÕES DE TELA
    MDIFormInicio.verificarPermissoesTelas

End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtIdOperador_GotFocus()

    With TxtIdOperador
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtIdOperador_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 115 Then ' F4
        'BtnCliente_Click
    ElseIf KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub TxtIdOperador_KeyPress(KeyAscii As Integer)
    
    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        If TxtIdOperador.Text = "" Then
            MsgBox ("Necessário informar um Operador!")
            KeyAscii = 0
        Else
            If STATUS = 1 Then
                ' FUNÇÃO QUE VERIFICA SE EXISTE ESTE OPERADOR NO SQL
                buscaOperador (TxtIdOperador.Text)
                KeyAscii = 0
            End If
        End If
    End If
End Sub



Public Function buscaOperador(ID As Integer)
    
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim rsDados As New ADODB.Recordset
    
    SQL = "SELECT IdOperador " & _
          "FROM Operador " & _
          "WHERE IdOperador = " & ID
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            
        Else
            MsgBox "Operador " & ID & " não encontrado"
            
            SQL = "DECLARE @maxID INT " & _
                  "SELECT @maxID = MAX(IdOperador) FROM Operador " & _
                  "SELECT @maxID AS Operador"
          
            rsDados.Open SQL, cn, adOpenStatic
                ID = rsDados("Operador")
            rsDados.Close
        End If
    rs.Close
    
    STATUS = 1
    preencheOperador (ID)
    
    SendKeys "{tab}"   ' Set the focus to the next control.
    ' VOLTA O FOCO PARA O ID CLIENTE
    SendKeys "+{tab}" ' SHIFT TAB
    
End Function

Private Function preencheOperador(ID As Integer)
    
    Dim rs As New ADODB.Recordset
    Dim rsDados As New ADODB.Recordset
    Dim SQL As String
    
    ' SE STATUS = 0 -> NOVO CADASTRO
    If STATUS = 0 Then
        
        ' HABILITO TODOS OS CAMPOS PARA PODER CADASTRAR O REGISTRO
        TxtIdOperador.Enabled = False
        TxtNome.Enabled = True
        TxtSenha.Enabled = True
        ChkVisualizarSenha.Enabled = True
        If FormLogin.ADMIN = 1 Then
            ChkADMIN.Enabled = True
            ChkInativo.Enabled = True
        Else
            ChkADMIN.Enabled = False
            ChkInativo.Enabled = False
        End If
        ChkFormCadastroOperador.Enabled = True
        ChkFormCadastroCategoria.Enabled = True
        ChkFormCadastroCliente.Enabled = True
        ChkFormCadastroProduto.Enabled = True
        ChkFormCadastroFormaPGTO.Enabled = True
        ChkFormTabelaUsuario.Enabled = True
        ChkFormPedidos.Enabled = True
        ChkFormEntrada.Enabled = True
        BtnNovo.Enabled = False
        BtnAlterar.Enabled = False
        BtnConfirmar.Enabled = True
        BtnCancelar.Enabled = True
        BtnAvancaRegistro.Enabled = False
        BtnVoltaRegistro.Enabled = False
        
        ' LIMPA TODOS OS CAMPOS
        TxtNome.Text = ""
        TxtSenha.Text = ""
        ChkVisualizarSenha.Value = Unchecked
        ChkADMIN.Value = Unchecked
        ChkInativo.Value = Unchecked
        ChkFormCadastroOperador.Value = Unchecked
        ChkFormCadastroCategoria.Value = Unchecked
        ChkFormCadastroCliente.Value = Unchecked
        ChkFormCadastroProduto.Value = Unchecked
        ChkFormCadastroFormaPGTO.Value = Unchecked
        ChkFormTabelaUsuario.Value = Unchecked
        ChkFormPedidos.Value = Unchecked
            
    ' SE STATUS CLIENTE = 1 -> CLIENTE JÁ CADASTRADO -> MOSTRA OS DADOS
    ElseIf STATUS = 1 Then
    
        SQL = "SELECT * " & _
              "FROM Operador " & _
              "WHERE IdOperador = " & ID
        
        rs.Open SQL, cn, adOpenStatic
            
            TxtIdOperador.Text = rs("IdOperador")
            TxtNome.Text = rs("nomeOperador")
            TxtSenha.Text = rs("senhaOperador")
            
            ' VERIFICA O INATIVO -> False = ATIVO | True = INATIVO
            If rs("statusOperador") = True Then
                ChkInativo = Checked
            Else
                ChkInativo = Unchecked
            End If
            
            ' VERIFICA O ADMINISTRADOR
            If rs("adminOperador") = True Then
                ChkADMIN = Checked
            Else
                ChkADMIN = Unchecked
            End If
            
            SQL = "SELECT * " & _
                  "FROM OperadorPermissaoTela " & _
                  "WHERE IdOperador = " & ID
            
            rsDados.Open SQL, cn, adOpenStatic
                ' VERIFICA A TELA CADASTRO DE OPERADOR
                If rsDados("CadastroOperador") = True Then
                    ChkFormCadastroOperador = Checked
                Else
                    ChkFormCadastroOperador = Unchecked
                End If
                
                ' VERIFICA A TELA CADASTRO DE CATEGORIA
                If rsDados("CadastroCategoria") = True Then
                    ChkFormCadastroCategoria = Checked
                Else
                    ChkFormCadastroCategoria = Unchecked
                End If
                
                ' VERIFICA A TELA CADASTRO DE CLIENTE
                If rsDados("CadastroCliente") = True Then
                    ChkFormCadastroCliente = Checked
                Else
                    ChkFormCadastroCliente = Unchecked
                End If
                
                ' VERIFICA A TELA CADASTRO PRODUTO
                If rsDados("CadastroProduto") = True Then
                    ChkFormCadastroProduto = Checked
                Else
                    ChkFormCadastroProduto = Unchecked
                End If
                
                ' VERIFICA A TELA CADASTRO PRODUTO
                If rsDados("CadastroFormaPGTO") = True Then
                    ChkFormCadastroFormaPGTO = Checked
                Else
                    ChkFormCadastroFormaPGTO = Unchecked
                End If
                
                ' VERIFICA A TELA TABELA DE CLIENTES
                If rsDados("TabelaUsuario") = True Then
                    ChkFormTabelaUsuario = Checked
                Else
                    ChkFormTabelaUsuario = Unchecked
                End If
                
                ' VERIFICA A TELA DE PEDIDOS
                If rsDados("Pedidos") = True Then
                    ChkFormPedidos = Checked
                Else
                    ChkFormPedidos = Unchecked
                End If
                
                ' VERIFICA A TELA DE PEDIDOS
                If rsDados("Entrada") = True Then
                    ChkFormEntrada.Value = Checked
                Else
                    ChkFormEntrada.Value = Unchecked
                End If
            rsDados.Close
        rs.Close
        
        Set rs = Nothing
        
        ' DESABILITO TODOS OS CAMPOS
        TxtIdOperador.Enabled = True
        TxtNome.Enabled = False
        TxtSenha.Enabled = False
        If FormLogin.ADMIN = 1 Then
            ChkVisualizarSenha.Enabled = True
        Else
            ChkVisualizarSenha.Enabled = False
        End If
        ChkADMIN.Enabled = False
        ChkInativo.Enabled = False
        ChkFormCadastroOperador.Enabled = False
        ChkFormCadastroCategoria.Enabled = False
        ChkFormCadastroCliente.Enabled = False
        ChkFormCadastroProduto.Enabled = False
        ChkFormCadastroFormaPGTO.Enabled = False
        ChkFormTabelaUsuario.Enabled = False
        ChkFormPedidos.Enabled = False
        ChkFormEntrada.Enabled = False
        BtnNovo.Enabled = True
        BtnCancelar.Enabled = True
        BtnConfirmar.Enabled = False
        BtnAlterar.Enabled = True
        BtnAvancaRegistro.Enabled = True
        BtnVoltaRegistro.Enabled = True
    
    ElseIf STATUS = 2 Then
        ' HABILITO TODOS OS CAMPOS PARA PODER ALTERAR O REGISTRO
        TxtIdOperador.Enabled = False
        TxtNome.Enabled = True
        TxtSenha.Enabled = True
        ChkVisualizarSenha.Enabled = True
        If FormLogin.ADMIN = 1 Then
            ChkADMIN.Enabled = True
            ChkInativo.Enabled = True
        Else
            ChkADMIN.Enabled = False
            ChkInativo.Enabled = False
        End If
        ChkFormCadastroOperador.Enabled = True
        ChkFormCadastroCategoria.Enabled = True
        ChkFormCadastroCliente.Enabled = True
        ChkFormCadastroProduto.Enabled = True
        ChkFormCadastroFormaPGTO.Enabled = True
        ChkFormTabelaUsuario.Enabled = True
        ChkFormPedidos.Enabled = True
        ChkFormEntrada.Enabled = True
        BtnAlterar.Enabled = False
        BtnNovo.Enabled = False
        BtnConfirmar.Enabled = True
        BtnCancelar.Enabled = True
        BtnAvancaRegistro.Enabled = False
        BtnVoltaRegistro.Enabled = False
    End If
    
End Function

Private Function finalizaForm()

    Dim validacao As Integer
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim Operador As String
    Dim ID As Integer
    
    If STATUS = 0 Or STATUS = 2 Then
        If STATUS = 0 Then
            validacao = MsgBox("Deseja cancelar o cadastro? Todas as informações serão perdidas!", vbYesNo)
            
            SQL = "DECLARE @maxID INT " & _
                  "SELECT @maxID = MAX(IdOperador) FROM Operador " & _
                  "SELECT @maxID AS Operador"
        ElseIf STATUS = 2 Then
            validacao = MsgBox("Deseja cancelar a alteraçao? Todas as alterações não serão salvas!", vbYesNo)
            
            SQL = "SELECT IdOperador as Operador FROM Operador WHERE IdOperador = " & TxtIdOperador.Text
        End If
        
        If validacao = vbYes Then
            rs.Open SQL, cn, adOpenStatic
                If IsNull(rs("Operador")) Then
                    ' DEFINE O STATUS PARA 1 -> NOVO REGISTRO
                    STATUS = 0
                    ID = 1
                    TxtIdOperador.Text = 1
                Else
                    ID = rs("Operador")
                    ' DEFINE O STATUS PARA 1 -> REGISTRO JÁ CADASTRADO
                    STATUS = 1
                End If
            rs.Close
            
            preencheOperador (ID)
            
        End If
    Else
        Unload Me
    End If
End Function

Private Sub TxtNome_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtNome_KeyPress(KeyAscii As Integer)
    
    'BLOQUEIA TODOS OS CARACTERES ESPECIAIS
    If (KeyAscii >= 33 And KeyAscii <= 39) Or (KeyAscii >= 40 And KeyAscii <= 44) Or (KeyAscii >= 46 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 94) Or (KeyAscii >= 123 And KeyAscii <= 126) Or (KeyAscii >= 162 And KeyAscii <= 163) Or KeyAscii = 95 Or KeyAscii = 168 Or KeyAscii = 172 Or KeyAscii = 176 Or (KeyAscii >= 178 And KeyAscii <= 180) Or KeyAscii = 185 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then ' ENTER
        SendKeys ("{tab}")
    End If
    
End Sub

Private Sub TxtSenha_KeyPress(KeyAscii As Integer)

    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then ' ENTER
        SendKeys ("{tab}")
    End If

End Sub
