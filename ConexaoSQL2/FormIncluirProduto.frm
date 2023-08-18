VERSION 5.00
Begin VB.Form FormCadastroCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro Cliente"
   ClientHeight    =   6375
   ClientLeft      =   7290
   ClientTop       =   4485
   ClientWidth     =   15075
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   15075
   Begin VB.Frame FrameTipo 
      Caption         =   "Tipo"
      Height          =   3135
      Left            =   13155
      TabIndex        =   33
      Top             =   2490
      Width           =   1710
      Begin VB.CheckBox ChkTipoFornecedor 
         Caption         =   "Fornecedor"
         Height          =   405
         Left            =   180
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox ChkTipoCliente 
         Caption         =   "Cliente"
         Height          =   405
         Left            =   180
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   300
         Width           =   1380
      End
   End
   Begin VB.ComboBox ComboBoxEstados 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3750
      Width           =   1080
   End
   Begin VB.CommandButton BtnAvancaRegistro 
      Caption         =   ">"
      Height          =   555
      Left            =   5445
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   450
      Width           =   675
   End
   Begin VB.CommandButton BtnVoltaRegistro 
      Caption         =   "<"
      Height          =   555
      Left            =   4680
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   450
      Width           =   675
   End
   Begin VB.ComboBox ComboBoxDocumento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7860
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame FrameMoradia 
      Caption         =   "Moradia"
      Height          =   1875
      Left            =   11205
      TabIndex        =   27
      Top             =   3645
      Width           =   1635
      Begin VB.OptionButton OptApartamento 
         Caption         =   "Apartamento"
         Height          =   420
         Left            =   195
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1110
         Width           =   1245
      End
      Begin VB.OptionButton OptCasa 
         Caption         =   "Casa"
         Height          =   495
         Left            =   195
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   375
         Width           =   795
      End
   End
   Begin VB.CommandButton BtnCliente 
      Caption         =   "->"
      Height          =   450
      Left            =   2100
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   495
      Width           =   675
   End
   Begin VB.CheckBox ChkInativo 
      Alignment       =   1  'Right Justify
      Caption         =   "INATIVO"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   13050
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1725
      Width           =   1785
   End
   Begin VB.TextBox TxtIdCliente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2970
      MaxLength       =   50
      TabIndex        =   1
      Top             =   465
      Width           =   1395
   End
   Begin VB.TextBox TxtBairro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   11
      Top             =   5520
      Width           =   3690
   End
   Begin VB.TextBox TxtNumero 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9870
      MaxLength       =   4
      TabIndex        =   10
      Top             =   4650
      Width           =   1005
   End
   Begin VB.TextBox TxtEndereco 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   9
      Top             =   4650
      Width           =   7560
   End
   Begin VB.TextBox TxtCidade 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4425
      MaxLength       =   50
      TabIndex        =   8
      Top             =   3750
      Width           =   4530
   End
   Begin VB.TextBox TxtCPF 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "###.###.###-##"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9285
      MaxLength       =   14
      TabIndex        =   4
      Top             =   1755
      Width           =   3480
   End
   Begin VB.TextBox TxtNascimento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9780
      MaxLength       =   10
      TabIndex        =   6
      Top             =   2760
      Width           =   2760
   End
   Begin VB.TextBox TxtEmail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1620
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2760
      Width           =   5670
   End
   Begin VB.TextBox TxtNome 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1590
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1755
      Width           =   5670
   End
   Begin VB.Frame Frame1 
      Height          =   1140
      Left            =   6630
      TabIndex        =   21
      Top             =   90
      Width           =   7815
      Begin VB.CommandButton BtnAlterar 
         Caption         =   "Alterar"
         Height          =   645
         Left            =   4320
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   285
         Width           =   1335
      End
      Begin VB.CommandButton BtnNovoCliente 
         Caption         =   "Novo"
         Height          =   645
         Left            =   435
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   285
         Width           =   1335
      End
      Begin VB.CommandButton BtnCancelar 
         Caption         =   "Cancelar"
         Height          =   645
         Left            =   6150
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   285
         Width           =   1335
      End
      Begin VB.CommandButton BtnConfirmarCadastro 
         Caption         =   "Confirmar"
         Height          =   645
         Left            =   2340
         TabIndex        =   13
         Top             =   285
         Width           =   1335
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   360
      TabIndex        =   32
      Top             =   3750
      Width           =   1110
   End
   Begin VB.Label Label10 
      Caption         =   "ID Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   420
      TabIndex        =   23
      Top             =   525
      Width           =   1650
   End
   Begin VB.Label Label9 
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9705
      TabIndex        =   20
      Top             =   4005
      Width           =   1365
   End
   Begin VB.Label Label8 
      Caption         =   "Endereço"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   150
      TabIndex        =   19
      Top             =   4650
      Width           =   1560
   End
   Begin VB.Label Label7 
      Caption         =   "Bairro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   705
      TabIndex        =   18
      Top             =   5550
      Width           =   1020
   End
   Begin VB.Label Label6 
      Caption         =   "Cidade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2955
      TabIndex        =   17
      Top             =   3750
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8235
      TabIndex        =   16
      Top             =   1275
      Width           =   1965
   End
   Begin VB.Label Label4 
      Caption         =   "Nascimento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7605
      TabIndex        =   15
      Top             =   2760
      Width           =   1965
   End
   Begin VB.Label Label3 
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   555
      TabIndex        =   14
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Label2 
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
      Height          =   435
      Left            =   480
      TabIndex        =   0
      Top             =   1755
      Width           =   990
   End
End
Attribute VB_Name = "FormCadastroCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public STATUS As Integer  ' 0 -> CLIETE NOVO / 1 -> CLIENTE JÁ FINALIZADO / 2 - EDIÇÃO DE PEDIDO JÁ FINALIZADO
Public Moradia As Integer ' 1 -> CASA / 2 -> APARTAMENTO
Public AnoAtual As Integer

Private Sub BtnAlterar_Click()

    ' ALTERA PARA O STATUS 2 -> ALTERAÇÃO DE PEDIDO JÁ FINALIZADO
    STATUS = 2
    ' CHAMA A FUNÇÃO QUE PREENCHE O PEDIDO
    preencheCliente (TxtIdCliente.Text)

End Sub

Private Sub BtnAlterar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnAvancaRegistro_Click()

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim codCliente As Integer
    
    SQL = "DECLARE @ID INT " & _
          "SET @ID = " & TxtIdCliente.Text & " " & _
          "SELECT TOP 1 IdCliente FROM Clientes WHERE IdCliente > @ID ORDER BY IdCliente"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            codCliente = rs("IdCliente")
        Else
            codCliente = TxtIdCliente.Text
        End If
    rs.Close
    
    TxtIdCliente.Text = codCliente
    preencheCliente (codCliente)

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

Private Sub BtnCliente_Click()
    FormBuscaCliente.FORMULARIO = "FormCadastroCliente"
    FormBuscaCliente.Show
    FormBuscaCliente.SetFocus
    MDIFormInicio.Enabled = False
End Sub

Private Sub BtnCliente_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnConfirmarCadastro_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then ' The ENTER key.
        SendKeys ("{TAB}")
    ElseIf KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnNovoCliente_Click()
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim numeroCliente As Integer
    
    SQL = "SELECT MAX(IdCliente)+1 AS Cliente FROM Clientes"
    
    rs.Open SQL, cn, adOpenStatic
        numeroCliente = rs("Cliente")
        TxtIdCliente.Text = rs("Cliente")
    rs.Close
    
    ' SETO O STATUS DA TELA PARA 0 -> CLIENTE NOVO
    STATUS = 0
    preencheCliente (numeroCliente)
End Sub

Private Sub BtnNovoCliente_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnVoltaItem_Click()

    
End Sub

Private Sub BtnVoltaRegistro_Click()

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim codCliente As Integer
    
    SQL = "DECLARE @ID INT " & _
          "SET @ID = " & TxtIdCliente.Text & " " & _
          "SELECT TOP 1 IdCliente FROM Clientes WHERE IdCliente < @ID ORDER BY IdCliente DESC"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            codCliente = rs("IdCliente")
        Else
            codCliente = TxtIdCliente.Text
        End If
    rs.Close
    
    TxtIdCliente.Text = codCliente
    preencheCliente (codCliente)


End Sub

Private Sub BtnVoltaRegistro_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub ChkInativo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then ' The ENTER key.
        SendKeys ("{TAB}")
    ElseIf KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub ChkTipoCliente_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub ChkTipoFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub ComboBoxDocumento_Click()
    
    If ComboBoxDocumento.Text = "CPF" Then
        If STATUS = 0 Then
            TxtCPF.Text = ""
        End If
        TxtCPF.MaxLength = 14
    ElseIf ComboBoxDocumento.Text = "CNPJ" Then
        If STATUS = 0 Then
            TxtCPF.Text = ""
        End If
        TxtCPF.MaxLength = 18
    End If

End Sub

Private Sub ComboBoxDocumento_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then ' The ENTER key.
        SendKeys ("{TAB}")
        KeyAscii = 0
    End If

End Sub

Private Sub ComboBoxEstados_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
    
End Sub

Private Sub ComboBoxEstados_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then ' The ENTER key.
        SendKeys ("{TAB}")
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Load()
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim ID As Integer
    
    SQL = "SELECT YEAR(getdate()) AS AnoAtual"
    
    ' PEGA O ANO ATUAL
    rs.Open SQL, cn, adOpenStatic
        AnoAtual = rs("AnoAtual")
    rs.Close
    
    Set rs = Nothing
    
    SQL = "DECLARE @maxCliente INT " & _
          "SELECT @maxCliente = MAX(IdCliente) FROM Clientes " & _
          "SELECT @maxCliente AS Cliente"
          
    ' ADICIONA OS ITENS NA COMBO BOX
    ComboBoxDocumento.AddItem ("CPF")
    ComboBoxDocumento.AddItem ("CNPJ")
    
    rs.Open SQL, cn, adOpenStatic
        If IsNull(rs("Cliente")) Then
            ' DEFINE O STATUS PARA 1 -> CLIENTE NOVO
            STATUS = 0
            ID = 1
            TxtIdCliente.Text = 1
        Else
            ID = rs("Cliente")
            ' DEFINE O STATUS PARA 1 -> CLIENTE JÁ CADASTRADO
            STATUS = 1
        End If
    rs.Close
    
    Set rs = Nothing
    
    SQL = "SELECT UF " & _
          "FROM Estados"
          
    ' ADICIONO OS ESTADOS NO COMBO BOX
    ComboBoxEstados.AddItem ("")
    rs.Open SQL, cn, adOpenStatic
        Do While rs.EOF = False
            ComboBoxEstados.AddItem (rs("UF"))
          rs.MoveNext
        Loop
    rs.Close
    
    ' DEFINE A MORADIA COMO
    Moradia = 0
    
    preencheCliente (ID)
    
End Sub

Private Sub BtnCancelar_Click()
    finalizaForm
End Sub

Private Sub BtnConfirmarCadastro_Click()

    Dim SQL As String
    Dim ativo As Integer
    Dim validacao As Integer
    Dim CMD As New ADODB.Command
    Dim rs As New ADODB.Recordset

    On Error GoTo TrataErro
    
    If TxtNome.Text = "" Then
        MsgBox "Preencha o campo Nome!"
        TxtNome.SetFocus
    ElseIf TxtCPF.Text = "" Then
        If ComboBoxDocumento.Text = "CPF" Then
            MsgBox "Preencha o campo CPF!"
        ElseIf ComboBoxDocumento.Text = "CNPJ" Then
            MsgBox "Preencha o campo CNPJ!"
        End If
        TxtCPF.SetFocus
    ElseIf TxtEmail.Text = "" Then
        MsgBox "Preencha o campo Email!"
        TxtEmail.SetFocus
    ElseIf TxtNascimento.Text = "" Then
        MsgBox "Preencha o campo Nascimento!"
        TxtNascimento.SetFocus
    ElseIf ComboBoxEstados.Text = "" Then
        MsgBox "Preencha o campo Estado!"
        ComboBoxEstados.SetFocus
    ElseIf TxtCidade.Text = "" Then
        MsgBox "Preencha o campo Cidade!"
        TxtCidade.SetFocus
    ElseIf TxtEndereco.Text = "" Then
        MsgBox "Preencha o campo Endereço!"
        TxtEndereco.SetFocus
    ElseIf TxtNumero.Text = "" Then
        MsgBox "Preencha o campo Número!"
        TxtNumero.SetFocus
    ElseIf TxtBairro.Text = "" Then
        MsgBox "Preencha o campo Bairro!"
        TxtBairro.SetFocus
    ElseIf OptCasa.Value = False And OptApartamento.Value = False Then
        MsgBox "Necessário selecionar o tipo de Moradia"
    ElseIf verificaTipoCliente = 1 Then
        MsgBox ("Necessário selecionar um tipo para o Cliente")
    Else
        ' SE ESTIVER TUDO COMPLETO EXECUTA
        
        If ChkInativo.Value = Checked Then
            ativo = 1
        Else
            ativo = 0
        End If
        
        ' SELECIONA A SP NO COMMAND
        If STATUS = 0 Then
            ' PASSA A CONEXÃO PARA O COMMAND
            CMD.ActiveConnection = cn
            CMD.CommandText = "adicionaCliente"
            CMD.CommandType = adCmdStoredProc
        ElseIf STATUS = 2 Then
            CMD.ActiveConnection = cn
            CMD.CommandText = "alteraCliente"
            CMD.CommandType = adCmdStoredProc
        End If
        
        ComboBoxEstados.Enabled = True
        
        ' PASSA OS PARAMETROS PARA O COMMAND
        CMD.Parameters.Append CMD.CreateParameter("RetornoOperacao", adInteger, adParamReturnValue, , 99)
        CMD.Parameters.Append CMD.CreateParameter("OUTPUT", adInteger, adParamOutput, , 99)
        CMD.Parameters.Append CMD.CreateParameter("CliNome", adVarChar, adParamInput, 50, TxtNome.Text)
        CMD.Parameters.Append CMD.CreateParameter("CPF", adVarChar, adParamInput, 14, TxtCPF.Text)
        CMD.Parameters.Append CMD.CreateParameter("EMAIL", adVarChar, adParamInput, 50, TxtEmail.Text)
        CMD.Parameters.Append CMD.CreateParameter("DtNasc", adVarChar, adParamInput, 10, TxtNascimento.Text)
        CMD.Parameters.Append CMD.CreateParameter("Estado", adVarChar, adParamInput, 50, ComboBoxEstados.Text)
        CMD.Parameters.Append CMD.CreateParameter("Cidade", adVarChar, adParamInput, 50, TxtCidade.Text)
        CMD.Parameters.Append CMD.CreateParameter("Bairro", adVarChar, adParamInput, 50, TxtBairro.Text)
        CMD.Parameters.Append CMD.CreateParameter("Endereco", adVarChar, adParamInput, 50, TxtEndereco.Text)
        CMD.Parameters.Append CMD.CreateParameter("Numero", adVarChar, adParamInput, 4, TxtNumero.Text)
        CMD.Parameters.Append CMD.CreateParameter("Status", adBoolean, adParamInput, , ativo)
        CMD.Parameters.Append CMD.CreateParameter("Moradia", adInteger, adParamInput, , Moradia)
        CMD.Parameters.Append CMD.CreateParameter("ID", adInteger, adParamInput, , TxtIdCliente.Text) ' UTILIZA APENAS PARA ATUALIZAR
        CMD.Parameters.Append CMD.CreateParameter("TipoCliente", adBoolean, adParamInput, , ChkTipoCliente.Value)
        CMD.Parameters.Append CMD.CreateParameter("TipoFornecedor", adBoolean, adParamInput, , ChkTipoFornecedor.Value)
        
        CMD.Execute
        
        validacao = CMD.Parameters("RetornoOperacao").Value
        
        If STATUS = 0 Then
            If validacao = 0 Then
                MsgBox ("Cliente cadastrado com sucesso!")
                
                ' FINALIZOU COM SUCESSO, PEGA O NÚMERO DO CLIENTE
                numeroCliente = TxtIdCliente.Text
                        
                ' DEFINE O STATUS PARA 1 -> CLIENTE JÁ CADASTRADO
                STATUS = 1
        
                preencheCliente (numeroCliente)
            ElseIf validacao = 1 Then
                MsgBox ("Ocorreu algum erro ao tentar cadastrar o cliente!")
            End If
        ElseIf STATUS = 2 Then
            If validacao = 0 Then
                ' ALTEROU O CLIENTE COM SUCESSO
                MsgBox ("Cliente alterado com sucesso!")
                
                ' FINALIZOU COM SUCESSO, PEGA O NÚMERO DO CLIENTE
                numeroCliente = TxtIdCliente.Text
                        
                ' DEFINE O STATUS PARA 1 -> CLIENTE JÁ CADASTRADO
                STATUS = 1
                
                preencheCliente (numeroCliente)
            ElseIf validacao = 1 Then
                MsgBox ("Ocorreu algum erro ao tentar alterar o cliente!")
            End If
        End If
    End If
    
    Exit Sub

TrataErro:
    MsgBox "Algum erro ocorreu ao carregar o Form - " & FormCadastroCliente.Name & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description
End Sub

Private Sub OptApartamento_Click()

    Moradia = 2

End Sub

Private Sub OptApartamento_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub OptCasa_Click()
    
    Moradia = 1
    
End Sub

Private Sub OptCasa_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtBairro_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtBairro_KeyPress(KeyAscii As Integer)

    'BLOQUEIA TODOS OS CARACTERES ESPECIAIS
    If (KeyAscii >= 33 And KeyAscii <= 39) Or (KeyAscii >= 40 And KeyAscii <= 44) Or (KeyAscii >= 46 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 94) Or (KeyAscii >= 123 And KeyAscii <= 126) Or (KeyAscii >= 162 And KeyAscii <= 163) Or KeyAscii = 95 Or KeyAscii = 168 Or KeyAscii = 172 Or KeyAscii = 176 Or (KeyAscii >= 178 And KeyAscii <= 180) Or KeyAscii = 185 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then ' The ENTER key.
        SendKeys ("{TAB}")
        KeyAscii = 0
    End If

End Sub

Private Sub TxtCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
    
End Sub

Private Sub TxtCidade_KeyPress(KeyAscii As Integer)
    
    'BLOQUEIA TODOS OS CARACTERES ESPECIAIS
    If (KeyAscii >= 33 And KeyAscii <= 39) Or (KeyAscii >= 40 And KeyAscii <= 44) Or (KeyAscii >= 46 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 94) Or (KeyAscii >= 123 And KeyAscii <= 126) Or (KeyAscii >= 162 And KeyAscii <= 163) Or KeyAscii = 95 Or KeyAscii = 168 Or KeyAscii = 172 Or KeyAscii = 176 Or (KeyAscii >= 178 And KeyAscii <= 180) Or KeyAscii = 185 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then ' The ENTER key.
        SendKeys ("{TAB}")
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtCPF_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode <> 8 Then
        If ComboBoxDocumento.Text = "CPF" Then
            If TxtCPF.SelStart = 3 Then
                TxtCPF.SelText = "."
            ElseIf TxtCPF.SelStart = 7 Then
                TxtCPF.SelText = "."
            ElseIf TxtCPF.SelStart = 11 Then
                TxtCPF.SelText = "-"
            End If
        ElseIf ComboBoxDocumento.Text = "CNPJ" Then
            If TxtCPF.SelStart = 2 Then
                TxtCPF.SelText = "."
            ElseIf TxtCPF.SelStart = 6 Then
                TxtCPF.SelText = "."
            ElseIf TxtCPF.SelStart = 10 Then
                TxtCPF.SelText = "/"
            ElseIf TxtCPF.SelStart = 15 Then
                TxtCPF.SelText = "-"
            End If
        End If
    End If
    
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
    
End Sub

Private Sub TxtCPF_KeyPress(KeyAscii As Integer)
    
    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then ' The ENTER key.
        SendKeys ("{TAB}")
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtCPF_LostFocus()

    Dim documento As String
    documento = Replace(TxtCPF.Text, ".", "")
    documento = Replace(documento, "-", "")
    documento = Replace(documento, "/", "")

    ' SÓ VERIFICA O CPF / CNPJ SE FOR STATUS 0 OU 2
    If STATUS = 0 Or STATUS = 2 Then
        If Len(TxtCPF.Text) > 0 Then
            If ComboBoxDocumento.Text = "CPF" Then
                If Not calculacpf(documento) Then
                    TxtCPF.Text = Format$(TxtCPF.Text, "@@@.@@@.@@@-@@")
                    MsgBox "CPF Inválido!"
                    TxtCPF.Text = ""
                    TxtCPF.SetFocus
                End If
            ElseIf ComboBoxDocumento.Text = "CNPJ" Then
                If Not ValidaCGC(documento) Then
                    TxtCPF.Text = Format$(TxtCPF.Text, "@@.@@@.@@@/@@@@-@@")
                    MsgBox "CNPJ Inválido!"
                    TxtCPF.Text = ""
                    TxtCPF.SetFocus
                End If
            End If
        End If
    ElseIf STATUS = 1 Then
    
    End If
End Sub

Private Sub TxtEmail_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
    
    'BLOQUEIA APENAS ANDERLINE E ASPAS
    If (KeyAscii = 39 Or KeyAscii = 34) Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then ' The ENTER key.
        SendKeys ("{TAB}")
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtEmail_LostFocus()

    If STATUS = 0 Or STATUS = 2 Then
        If TxtEmail.Text = "" Then
        
        Else
            If InStr(TxtEmail.Text, "@") = 0 Or InStr(TxtEmail.Text, ".") = 0 Then
                MsgBox ("Email inválido")
                TxtEmail.SetFocus
            Else
                
            End If
        End If
    ElseIf STATUS = 1 Then
        
    End If
    
End Sub

Private Sub TxtEndereco_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtEndereco_KeyPress(KeyAscii As Integer)

    'BLOQUEIA TODOS OS CARACTERES ESPECIAIS
    If (KeyAscii >= 33 And KeyAscii <= 39) Or (KeyAscii >= 40 And KeyAscii <= 44) Or (KeyAscii >= 46 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 94) Or (KeyAscii >= 123 And KeyAscii <= 126) Or (KeyAscii >= 162 And KeyAscii <= 163) Or KeyAscii = 95 Or KeyAscii = 168 Or KeyAscii = 172 Or KeyAscii = 176 Or (KeyAscii >= 178 And KeyAscii <= 180) Or KeyAscii = 185 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If

    If KeyAscii = 13 Then ' The ENTER key.
        SendKeys ("{TAB}")
        KeyAscii = 0
    End If

End Sub

Private Sub TxtEstado_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then ' The ENTER key.
        SendKeys ("{TAB}")
    ElseIf KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtEstado_KeyPress(KeyAscii As Integer)

    'BLOQUEIA TODOS OS CARACTERES ESPECIAIS
    If (KeyAscii >= 33 And KeyAscii <= 39) Or (KeyAscii >= 40 And KeyAscii <= 44) Or (KeyAscii >= 46 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 94) Or (KeyAscii >= 123 And KeyAscii <= 126) Or (KeyAscii >= 162 And KeyAscii <= 163) Or KeyAscii = 95 Or KeyAscii = 168 Or KeyAscii = 172 Or KeyAscii = 176 Or (KeyAscii >= 178 And KeyAscii <= 180) Or KeyAscii = 185 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If

End Sub

Private Sub TxtIdCliente_GotFocus()

    With TxtIdCliente
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtIdCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 115 Then ' F4
        BtnCliente_Click
    ElseIf KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub TxtIdCliente_KeyPress(KeyAscii As Integer)
    
    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        If TxtIdCliente.Text = "" Then
            MsgBox ("Necessário informar um Cliente!")
            KeyAscii = 0
        Else
            If STATUS = 1 Then
                ' FUNÇÃO QUE VERIFICA SE EXISTE ESTE PEDIDO NO SQL
                buscaCliente (TxtIdCliente.Text)
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub TxtNascimento_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
    
End Sub

Public Function buscaCliente(numeroCliente As Integer)
    
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim rsDados As New ADODB.Recordset
    
    SQL = "SELECT IdCliente " & _
          "FROM Clientes " & _
          "WHERE IdCliente = " & numeroCliente
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            
        Else
            MsgBox "Cliente " & numeroCliente & " não encontrado"
            
            SQL = "DECLARE @maxCliente INT " & _
                  "SELECT @maxCliente = MAX(IdCliente) FROM Clientes " & _
                  "SELECT @maxCliente AS Cliente"
          
            rsDados.Open SQL, cn, adOpenStatic
                numeroCliente = rsDados("Cliente")
            rsDados.Close
        End If
    rs.Close
    
    STATUS = 1
    preencheCliente (numeroCliente)
    
    SendKeys "{tab}"   ' Set the focus to the next control.
    ' VOLTA O FOCO PARA O ID CLIENTE
    SendKeys "+{tab}" ' SHIFT TAB
    
End Function

Private Function preencheCliente(numeroCliente As Integer)
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    ' SE STATUS = 0 -> NOVO CADASTRO
    If STATUS = 0 Then
        
        ' HABILITO TODOS OS CAMPOS PARA PODER CADASTRAR O NOVO CLIENTE
        TxtIdCliente.Enabled = False
        TxtNome.Enabled = True
        TxtCPF.Enabled = True
        TxtEmail.Enabled = True
        TxtNascimento.Enabled = True
        ComboBoxEstados.Enabled = True
        TxtCidade.Enabled = True
        TxtEndereco.Enabled = True
        TxtNumero.Enabled = True
        TxtBairro.Enabled = True
        If FormLogin.ADMIN = 1 Then
            ChkInativo.Enabled = True
        End If
        OptApartamento.Enabled = True
        OptCasa.Enabled = True
        BtnNovoCliente.Enabled = False
        BtnConfirmarCadastro.Enabled = True
        BtnCancelar.Enabled = True
        BtnAlterar.Enabled = False
        ComboBoxDocumento.Enabled = True
        BtnAvancaRegistro.Enabled = False
        BtnVoltaRegistro.Enabled = False
        ChkTipoCliente.Enabled = True
        ChkTipoFornecedor.Enabled = True
        
        ' LIMPA TODOS OS CAMPOS
        TxtNome.Text = ""
        TxtCPF.Text = ""
        TxtEmail.Text = ""
        TxtNascimento.Text = ""
        TxtCidade.Text = ""
        TxtEndereco.Text = ""
        TxtNumero.Text = ""
        TxtBairro.Text = ""
        ChkInativo.Value = Unchecked
        OptApartamento.Value = False
        OptCasa.Value = False
        ComboBoxDocumento.Text = "CPF"
        ChkTipoCliente.Value = Unchecked
        ChkTipoFornecedor.Value = Unchecked
        
    ' SE STATUS CLIENTE = 1 -> CLIENTE JÁ CADASTRADO -> MOSTRA OS DADOS
    ElseIf STATUS = 1 Then
    
        SQL = "SELECT *, CONVERT(VARCHAR(10), CONVERT(DATE, CliDtNascimento, 126), 103) AS DataConvertida " & _
              "FROM Clientes " & _
              "WHERE IdCliente = " & numeroCliente
        
        rs.Open SQL, cn, adOpenStatic
            TxtIdCliente.Text = rs("IdCliente")
            TxtNome.Text = rs("CliNome")
            ' VERIFICA SE É CPF OU CNPJ
            If Len(rs("CliCPF")) = 14 Then
                ComboBoxDocumento.Text = "CPF"
            ElseIf Len(rs("CliCPF")) = 18 Then
                ComboBoxDocumento.Text = "CNPJ"
            End If
            TxtCPF.Text = rs("CliCPF")
            TxtEmail.Text = rs("CliEmail")
            TxtNascimento.Text = rs("DataConvertida")
            'TxtEstado.Text = rs("CliEstado")
            ComboBoxEstados.Text = rs("CliEstado")
            TxtCidade.Text = rs("CliCidade")
            TxtEndereco.Text = rs("CliEndereco")
            TxtNumero.Text = rs("CliNumero")
            TxtBairro.Text = rs("CliBairro")
            ' VERIFICA SE O CLIENTE ESTÁ INATIVO -> False = ATIVO | True = INATIVO
            If rs("CliStatus") = True Then
                ChkInativo.Value = Checked
            Else
                ChkInativo.Value = Unchecked
            End If
            ' VERIFICA A MORADIA -> 1 = CASA | 2 = APARTAMENTO
            If rs("Moradia") = 1 Then
                OptCasa.Value = True
            ElseIf rs("Moradia") = 2 Then
                OptApartamento.Value = True
            End If
        rs.Close
        Set rs = Nothing
        
        ' BUSCO O TIPO CLIENTE
        
        SQL = "SELECT * " & _
              "FROM TipoClientes " & _
              "WHERE IdCliente = " & numeroCliente
        
        rs.Open SQL, cn, adOpenStatic
            If rs("TipoCliente") = False Then
                ChkTipoCliente.Value = Unchecked
            Else
                ChkTipoCliente.Value = Checked
            End If
                
            If rs("TipoFornecedor") = False Then
                ChkTipoFornecedor.Value = Unchecked
            Else
                ChkTipoFornecedor.Value = Checked
            End If
        rs.Close
        
        ' DESABILITO TODOS OS CAMPOS
        TxtIdCliente.Enabled = True
        TxtNome.Enabled = False
        TxtCPF.Enabled = False
        TxtEmail.Enabled = False
        TxtNascimento.Enabled = False
        'TxtEstado.Enabled = False
        ComboBoxEstados.Enabled = False
        TxtCidade.Enabled = False
        TxtEndereco.Enabled = False
        TxtNumero.Enabled = False
        TxtBairro.Enabled = False
        OptApartamento.Enabled = False
        OptCasa.Enabled = False
        BtnNovoCliente.Enabled = True
        BtnConfirmarCadastro.Enabled = False
        BtnCancelar.Enabled = False
        BtnAlterar.Enabled = True
        ComboBoxDocumento.Enabled = False
        ChkInativo.Enabled = False
        BtnAvancaRegistro.Enabled = True
        BtnVoltaRegistro.Enabled = True
        ChkTipoCliente.Enabled = False
        ChkTipoFornecedor.Enabled = False
    
    ElseIf STATUS = 2 Then
        ' HABILITO TODOS OS CAMPOS PARA PODER CADASTRAR O NOVO CLIENTE
        TxtIdCliente.Enabled = False
        ' VERIFICA SE É CPF OU CNPJ
        If Len(TxtCPF.Text) = 14 Then
            ComboBoxDocumento.Text = "CPF"
        ElseIf Len(TxtCPF.Text) = 18 Then
            ComboBoxDocumento.Text = "CNPJ"
        End If
        TxtCPF.Enabled = True
        TxtEmail.Enabled = True
        TxtNascimento.Enabled = True
        'TxtEstado.Enabled = True
        ComboBoxEstados.Enabled = True
        TxtCidade.Enabled = True
        TxtEndereco.Enabled = True
        TxtNumero.Enabled = True
        TxtBairro.Enabled = True
        If FormLogin.ADMIN = 1 Then
            ChkInativo.Enabled = True
        End If
        OptApartamento.Enabled = True
        OptCasa.Enabled = True
        BtnNovoCliente.Enabled = False
        BtnConfirmarCadastro.Enabled = True
        BtnCancelar.Enabled = True
        BtnAlterar.Enabled = False
        ComboBoxDocumento.Enabled = True
        ChkTipoCliente.Enabled = True
        ChkTipoFornecedor.Enabled = True
    End If
    
End Function

Private Sub TxtNascimento_KeyPress(KeyAscii As Integer)

    Dim tamanho As Integer

    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
    If KeyAscii <> 8 Then
        If TxtNascimento.SelStart = 2 Then
            TxtNascimento.SelText = "/"
        ElseIf TxtNascimento.SelStart = 5 Then
            TxtNascimento.SelText = "/"
        End If
    End If
    
    ' ADICIONA O TAMANHO DO TXT NA VARIAVEL TAMANHO
    tamanho = Len(TxtNascimento.Text)
    
    If KeyAscii = 13 Then ' The ENTER key.
        If TxtNascimento.Text = "" Then
            SendKeys ("{TAB}")
            KeyAscii = 0
        ElseIf tamanho <> 10 Then ' VERIFICA O TAMANHO
            MsgBox ("Data com tamanho incorreto!")
            KeyAscii = 0
        Else
            SendKeys ("{TAB}")
            KeyAscii = 0
        End If
    End If
End Sub

Private Function finalizaForm()

    Dim validacao As Integer
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim Cliente As String
    
    If STATUS = 0 Or STATUS = 2 Then
        If STATUS = 0 Then
            validacao = MsgBox("Deseja cancelar o cadastro? Todas as informações serão perdidas!", vbYesNo)
            
            SQL = "DECLARE @maxCliente INT " & _
                  "SELECT @maxCliente = MAX(IdCliente) FROM Clientes " & _
                  "SELECT @maxCliente AS Cliente"
        ElseIf STATUS = 2 Then
            validacao = MsgBox("Deseja cancelar a alteraçao? Todas as alterações não serão salvas!", vbYesNo)
            
            SQL = "SELECT IdCliente as Cliente FROM Clientes WHERE IdCliente = " & TxtIdCliente.Text
        End If
        
        If validacao = vbYes Then
            rs.Open SQL, cn, adOpenStatic
                If IsNull(rs("Cliente")) Then
                    ' DEFINE O STATUS PARA 1 -> CLIENTE NOVO
                    STATUS = 0
                    numeroCliente = 1
                    TxtIdCliente.Text = 1
                Else
                    numeroCliente = rs("Cliente")
                    ' DEFINE O STATUS PARA 1 -> CLIENTE JÁ CADASTRADO
                    STATUS = 1
                End If
            rs.Close
            
            ' DEFINE A MORADIA COMO 0
            Moradia = 0
            
            preencheCliente (numeroCliente)
        End If
    Else
        Unload Me
    End If
End Function

Private Sub TxtNascimento_LostFocus()

    If STATUS = 0 Or STATUS = 2 Then
        If TxtNascimento.Text = "" Then
        
        Else
            verificarData
        End If
    ElseIf STATUS = 1 Then
        
    End If

End Sub

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
    
    If KeyAscii = 13 Then ' The ENTER key.
        SendKeys ("{TAB}")
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
    
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)

    'BLOQUEIA TODOS OS CARACTERES ESPECIAIS
    If (KeyAscii >= 33 And KeyAscii <= 39) Or (KeyAscii >= 40 And KeyAscii <= 44) Or (KeyAscii >= 46 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 94) Or (KeyAscii >= 123 And KeyAscii <= 126) Or (KeyAscii >= 162 And KeyAscii <= 163) Or KeyAscii = 95 Or KeyAscii = 168 Or KeyAscii = 172 Or KeyAscii = 176 Or (KeyAscii >= 178 And KeyAscii <= 180) Or KeyAscii = 185 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then ' The ENTER key.
        SendKeys ("{TAB}")
        KeyAscii = 0
    End If

End Sub

Function calculacpf(CPF As String) As Boolean

    On Error GoTo Err_CPF
    Dim I As Integer 'utilizada nos FOR... NEXT
    Dim strcampo As String 'armazena do CPF que será utilizada para o cálculo
    Dim strCaracter As String 'armazena os digitos do CPF da direita para a esquerda
    Dim intNumero As Integer 'armazena o digito separado para cálculo (uma a um)
    Dim intMais As Integer 'armazena o digito específico multiplicado pela sua base
    Dim lngSoma As Long 'armazena a soma dos digitos multiplicados pela sua base(intmais)
    Dim dblDivisao As Double 'armazena a divisão dos digitos*base por 11
    Dim lngInteiro As Long 'armazena inteiro da divisão
    Dim intResto As Integer 'armazena o resto
    Dim intDig1 As Integer 'armazena o 1º digito verificador
    Dim intDig2 As Integer 'armazena o 2º digito verificador
    Dim strConf As String 'armazena o digito verificador
    
    lngSoma = 0
    intNumero = 0
    intMais = 0
    strcampo = Left(CPF, 9)
    
    'Inicia cálculos do 1º dígito
    For I = 2 To 10
        strCaracter = Right(strcampo, I - 1)
        intNumero = Left(strCaracter, 1)
        intMais = intNumero * I
        lngSoma = lngSoma + intMais
    Next I
    dblDivisao = lngSoma / 11
    
    lngInteiro = Int(dblDivisao) * 11
    intResto = lngSoma - lngInteiro
    If intResto = 0 Or intResto = 1 Then
        intDig1 = 0
    Else
        intDig1 = 11 - intResto
    End If
    
    strcampo = strcampo & intDig1 'concatena o CPF com o primeiro digito verificador
    lngSoma = 0
    intNumero = 0
    intMais = 0
    'Inicia cálculos do 2º dígito
    For I = 2 To 11
        strCaracter = Right(strcampo, I - 1)
        intNumero = Left(strCaracter, 1)
        intMais = intNumero * I
        lngSoma = lngSoma + intMais
    Next I
    dblDivisao = lngSoma / 11
    lngInteiro = Int(dblDivisao) * 11
    intResto = lngSoma - lngInteiro
    If intResto = 0 Or intResto = 1 Then
        intDig2 = 0
    Else
        intDig2 = 11 - intResto
    End If
    strConf = intDig1 & intDig2
    
    'Caso o CPF esteja errado dispara a mensagem


    If TxtCPF.Text = "111.111.111-11" Then
        calculacpf = False
        Exit Function
    ElseIf TxtCPF.Text = "222.222.222-22" Then
        calculacpf = False
        Exit Function
    ElseIf TxtCPF.Text = "333.333.333-33" Then
        calculacpf = False
        Exit Function
    ElseIf TxtCPF.Text = "444.444.444-44" Then
        calculacpf = False
        Exit Function
    ElseIf TxtCPF.Text = "555.555.555-55" Then
        calculacpf = False
        Exit Function
    ElseIf TxtCPF.Text = "666.666.666-66" Then
        calculacpf = False
        Exit Function
    ElseIf TxtCPF.Text = "777.777.777-77" Then
        calculacpf = False
        Exit Function
    ElseIf TxtCPF.Text = "888.888.888-88" Then
        calculacpf = False
        Exit Function
    ElseIf TxtCPF.Text = "999.999.999-99" Then
        calculacpf = False
        Exit Function
    End If

    If strConf <> Right(CPF, 2) Then
        calculacpf = False
    Else
        calculacpf = True
    End If
    Exit Function
    
Exit_CPF:
        Exit Function
Err_CPF:
        MsgBox Error$
        Resume Exit_CPF
End Function

Public Function CalculaCGC(Numero As String) As String

    Dim I As Integer
    Dim prod As Integer
    Dim mult As Integer
    Dim digito As Integer
    
    If Not IsNumeric(Numero) Then
       CalculaCGC = ""
       Exit Function
    End If
    
    mult = 2
    For I = Len(Numero) To 1 Step -1
      prod = prod + Val(Mid(Numero, I, 1)) * mult
      mult = IIf(mult = 9, 2, mult + 1)
    Next
    
    digito = 11 - Int(prod Mod 11)
    digito = IIf(digito = 10 Or digito = 11, 0, digito)
    
    CalculaCGC = Trim(Str(digito))

End Function

Public Function ValidaCGC(CGC As String) As Boolean

    If CalculaCGC(Left(CGC, 12)) <> Mid(CGC, 13, 1) Then
       ValidaCGC = False
       Exit Function
    End If
    
    If CalculaCGC(Left(CGC, 13)) <> Mid(CGC, 14, 1) Then
       ValidaCGC = False
       Exit Function
    End If
    
    ValidaCGC = True

End Function

Private Function verificarData()

    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim validacao As Integer
    
    SQL = "SET DATEFORMAT dmy; " & _
          "SELECT ISDATE('" & TxtNascimento.Text & "') AS Validacao"
          
    rs.Open SQL, cn, adOpenStatic
        validacao = rs("Validacao")
    rs.Close
    
    If validacao = 1 Then
        ' DATA É VALIDA
        If Mid(TxtNascimento.Text, 7, 10) >= AnoAtual Then
            MsgBox ("Ano de Nascimento Inválido!")
            TxtNascimento.Text = ""
            TxtNascimento.SetFocus
        Else
            ' ANO VÁLIDO
        End If
    ElseIf validacao = 0 Then
        MsgBox ("Data Inválida!")
        TxtNascimento.Text = ""
        TxtNascimento.SetFocus
    End If

End Function

Private Function verificaTipoCliente()
    
    ' FUNÇÃO PARA VALIDAR SE TEM ALGUM TIPO DE CLIENTE SELECIONADO
    
    Dim validacao As Integer
    
    validacao = 0

    If ChkTipoCliente.Value = Unchecked And ChkTipoFornecedor.Value = Unchecked Then
        validacao = 1
    End If

    verificaTipoCliente = validacao

End Function
