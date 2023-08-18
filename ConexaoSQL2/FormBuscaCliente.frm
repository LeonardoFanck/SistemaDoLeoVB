VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FormBuscaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busca Cliente"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9990
   Begin VB.TextBox TxtDados 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2070
      TabIndex        =   1
      Top             =   135
      Width           =   5265
   End
   Begin VB.CommandButton BtnMostrarTodos 
      Caption         =   "Mostrar Todos"
      Height          =   510
      Left            =   7530
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   2160
   End
   Begin VB.ComboBox ComboSelecao 
      Height          =   315
      ItemData        =   "FormBuscaCliente.frx":0000
      Left            =   180
      List            =   "FormBuscaCliente.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   195
      Width           =   1800
   End
   Begin MSComctlLib.ListView ListViewListagemClientes 
      Height          =   4395
      Left            =   270
      TabIndex        =   2
      Top             =   855
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   7752
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   1191
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   4154
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "CPF"
         Object.Width           =   2170
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Email"
         Object.Width           =   4339
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Data Nasc"
         Object.Width           =   1879
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cidade"
         Object.Width           =   2778
      EndProperty
   End
End
Attribute VB_Name = "FormBuscaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tipoBusca As String
Public FORMULARIO As String

Private Sub BtnPesquisar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub BtnMostrarTodos_Click()
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    TxtDados.Text = ""
    
    If FORMULARIO = "FormPedidos" Then
        SQL = "SELECT TOP 15 Cli.IdCliente AS Codigo, Cli.CliNome AS Nome, Cli.CliCPF AS CPF, Cli.CliEmail AS Email, convert(varchar, Cli.CliDtNascimento, 103) AS DtNasc, Cli.CliCidade AS Cidade " & _
              "FROM Clientes AS Cli " & _
              "JOIN TipoClientes AS Tipo ON Cli.IdCliente = Tipo.IdCliente " & _
              "WHERE Cli.CliStatus = 0 AND Tipo.TipoCliente = 1 " & _
              "ORDER BY Codigo"
    ElseIf FORMULARIO = "FormEntrada" Then
        SQL = "SELECT TOP 15 Cli.IdCliente AS Codigo, Cli.CliNome AS Nome, Cli.CliCPF AS CPF, Cli.CliEmail AS Email, convert(varchar, Cli.CliDtNascimento, 103) AS DtNasc, Cli.CliCidade AS Cidade " & _
              "FROM Clientes AS Cli " & _
              "JOIN TipoClientes AS Tipo ON Cli.IdCliente = Tipo.IdCliente " & _
              "WHERE Cli.CliStatus = 0 AND Tipo.TipoFornecedor = 1 " & _
              "ORDER BY Codigo"
    ElseIf FORMULARIO = "FormCadastroCliente" Then
        SQL = "SELECT IdCliente AS Codigo, CliNome AS Nome, CliCPF AS CPF, CliEmail AS Email, convert(varchar, CliDtNascimento, 103) AS DtNasc, CliCidade AS Cidade " & _
              "FROM Clientes " & _
              "ORDER BY Codigo"
    End If
          
    ' LIMPA A LISTA SEMPRE
    ListViewListagemClientes.ListItems.Clear
    
    rs.Open SQL, cn, adOpenStatic
            If rs.EOF = True Then
                
            Else
                Do While rs.EOF = False
                    Set Cliente = ListViewListagemClientes.ListItems.Add(, , rs("Codigo"))
                    Cliente.SubItems(1) = (rs("Nome"))
                    Cliente.SubItems(2) = (rs("CPF"))
                    Cliente.SubItems(3) = (rs("Email"))
                    Cliente.SubItems(4) = (rs("DtNasc"))
                    Cliente.SubItems(5) = (rs("Cidade"))
                  'SE MOVE PARA O PROXIMO REGISTRO
                  rs.MoveNext
                Loop
            End If
        rs.Close
End Sub

Private Sub BtnMostrarTodos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub ComboSelecao_Click()
    If ComboSelecao.Text = "Codigo" Then
        TxtDados.Text = ""
        tipoBusca = "IdCliente"
    ElseIf ComboSelecao.Text = "Nome" Then
        TxtDados.Text = ""
        tipoBusca = "CliNome"
    ElseIf ComboSelecao.Text = "CPF" Then
        TxtDados.Text = ""
        tipoBusca = "CliCPF"
    ElseIf ComboSelecao.Text = "Email" Then
        TxtDados.Text = ""
        tipoBusca = "CliEmail"
    ElseIf ComboSelecao.Text = "Data Nasc" Then
        TxtDados.Text = ""
        tipoBusca = "CliDtNascimento"
    ElseIf ComboSelecao.Text = "Cidade" Then
        TxtDados.Text = ""
        tipoBusca = "CliCidade"
    End If
End Sub

Private Sub ComboSelecao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Sub Form_Load()
    
    'ADICIONANDO ITENS NO COMBO BOX
    ComboSelecao.AddItem ("Codigo")
    ComboSelecao.AddItem ("Nome")
    ComboSelecao.AddItem ("CPF")
    ComboSelecao.AddItem ("Email")
    ComboSelecao.AddItem ("Data Nasc")
    ComboSelecao.AddItem ("Cidade")
    
    'VALOR PADRÃO AO ABRIR A TELA
    ComboSelecao.Text = "Nome"
    'SETA O VALOR PADRÃO DE BUSCA
    tipoBusca = "CliNome"
    
    ' EXECUTO A FUNÇÃO QUE MOSTRA OS 10 PRIMIROS CLIENTES NO GRID
    primeirosDezClientes
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' ATIVA NOVAMENTE O MDI
    MDIFormInicio.Enabled = True
    
    If FORMULARIO = "FormPedidos" Then
        If FormPedidos.STATUS = 0 Then
            FormPedidos.TxtCliente.SetFocus
        End If
    ElseIf FORMULARIO = "FormEntrada" Then
        If FormEntrada.STATUS = 0 Then
            FormEntrada.TxtFornecedor.SetFocus
        End If
    End If
End Sub

Private Sub ListViewListagemClientes_DblClick()
    
    Dim IDCliente As Integer
    
    ' PEGA O ID DO CLENTE DA LISTA
    IDCliente = ListViewListagemClientes.SelectedItem.Text
    
    ' CHAMA A FUNCAO VERIFICA QUAL FORM CHAMOU A TELA
    VerificaForm (IDCliente)
    
End Sub

Private Sub ListViewListagemClientes_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim IDCliente As Integer
    
    If KeyCode = 27 Then ' ESC
        finalizaForm
    ElseIf KeyCode = 13 Then ' ENTER
        ' PEGA O ID DO CLENTE DA LISTA
        IDCliente = ListViewListagemClientes.SelectedItem.Text

        ' CHAMA A FUNCAO VERIFICA QUAL FORM CHAMOU A TELA
        VerificaForm (IDCliente)
    End If
End Sub

Private Function VerificaForm(IDCliente As Integer)
    
    ' VERIFICA QUAL FORM CHAMOU A TELA
    If FORMULARIO = "FormPedidos" Then
        If FormPedidos.STATUS = 0 Then
            FormPedidos.TxtCliente.Text = IDCliente
            FormPedidos.PegaNomeCliente
            finalizaForm
            FormPedidos.SetFocus
        End If
    ElseIf FORMULARIO = "FormCadastroCliente" Then
        If FormCadastroCliente.STATUS = 1 Then
            FormCadastroCliente.TxtIdCliente.Text = IDCliente
            FormCadastroCliente.buscaCliente (IDCliente) ' ENVIA O ID DO CLIENTE PARA A FUNÇÃO QUE VERIFICA SE EXISTE
            finalizaForm
            FormCadastroCliente.SetFocus
        End If
    ElseIf FORMULARIO = "FormEntrada" Then
        If FormEntrada.STATUS = 0 Then
            FormEntrada.TxtFornecedor.Text = IDCliente
            If FormEntrada.PegaNomeFornecedor = 0 Then
                Exit Function
            End If
            finalizaForm
            FormEntrada.SetFocus
        End If
    End If
End Function

Private Sub TxtDados_Change()
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim SQL2 As String
    
    
    If FORMULARIO = "FormCadastroCliente" Then
        SQL = "SELECT IdCliente AS Codigo, CliNome AS Nome, CliCPF AS CPF, CliEmail AS Email, convert(varchar, CliDtNascimento, 103) AS DtNasc, CliCidade AS Cidade " & _
              "FROM Clientes " & _
              "WHERE " & tipoBusca & " LIKE '%" & TxtDados.Text & "%' " & _
              "ORDER BY " & tipoBusca
    ElseIf FORMULARIO = "FormPedidos" Then
        SQL = "SELECT TOP 15 Cli.IdCliente AS Codigo, Cli.CliNome AS Nome, Cli.CliCPF AS CPF, Cli.CliEmail AS Email, convert(varchar, Cli.CliDtNascimento, 103) AS DtNasc, Cli.CliCidade AS Cidade " & _
              "FROM Clientes AS Cli " & _
              "JOIN TipoClientes AS Tipo ON Cli.IdCliente = Tipo.IdCliente " & _
              "WHERE " & tipoBusca & " LIKE '%" & TxtDados.Text & "%' AND Cli.CliStatus = 0 AND Tipo.TipoCliente = 1 " & _
              "ORDER BY Codigo"
    ElseIf FORMULARIO = "FormEntrada" Then
        SQL = "SELECT TOP 15 Cli.IdCliente AS Codigo, Cli.CliNome AS Nome, Cli.CliCPF AS CPF, Cli.CliEmail AS Email, convert(varchar, Cli.CliDtNascimento, 103) AS DtNasc, Cli.CliCidade AS Cidade " & _
              "FROM Clientes AS Cli " & _
              "JOIN TipoClientes AS Tipo ON Cli.IdCliente = Tipo.IdCliente " & _
              "WHERE " & tipoBusca & " LIKE '%" & TxtDados.Text & "%' AND Cli.CliStatus = 0 AND Tipo.TipoFornecedor = 1 " & _
              "ORDER BY Codigo"
    End If
    
    'SEMPRE LIMPA O LISTVIEW
    ListViewListagemClientes.ListItems.Clear
    
    If TxtDados.Text = "" Then
        'SE ESTIVER VAZIO A PESQUISA, BUSCA OS 10 PRIMEIROS CLIENTE DO SELECT
        primeirosDezClientes
    Else
        rs.Open SQL, cn, adOpenStatic
            If rs.EOF = True Then
                
            Else
                Do While rs.EOF = False
                    Set Cliente = ListViewListagemClientes.ListItems.Add(, , rs("Codigo"))
                    Cliente.SubItems(1) = (rs("Nome"))
                    Cliente.SubItems(2) = (rs("CPF"))
                    Cliente.SubItems(3) = (rs("Email"))
                    Cliente.SubItems(4) = (rs("DtNasc"))
                    Cliente.SubItems(5) = (rs("Cidade"))
                  'SE MOVE PARA O PROXIMO REGISTRO
                  rs.MoveNext
                Loop
            End If
        rs.Close
    End If
End Sub

Private Sub TxtDados_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Function primeirosDezClientes()

    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    If FORMULARIO = "FormCadastroCliente" Then
        SQL = "SELECT TOP 15 IdCliente AS Codigo, CliNome AS Nome, CliCPF AS CPF, CliEmail AS Email, convert(varchar, CliDtNascimento, 103) AS DtNasc, CliCidade AS Cidade " & _
              "FROM Clientes " & _
              "WHERE " & tipoBusca & " LIKE '%" & TxtDados.Text & "%' " & _
              "ORDER BY Codigo"
    ElseIf FORMULARIO = "FormPedidos" Then
        SQL = "SELECT TOP 15 Cli.IdCliente AS Codigo, Cli.CliNome AS Nome, Cli.CliCPF AS CPF, Cli.CliEmail AS Email, convert(varchar, Cli.CliDtNascimento, 103) AS DtNasc, Cli.CliCidade AS Cidade " & _
              "FROM Clientes AS Cli " & _
              "JOIN TipoClientes AS Tipo ON Cli.IdCliente = Tipo.IdCliente " & _
              "WHERE " & tipoBusca & " LIKE '%" & TxtDados.Text & "%' AND Cli.CliStatus = 0 AND Tipo.TipoCliente = 1 " & _
              "ORDER BY Codigo"
    ElseIf FORMULARIO = "FormEntrada" Then
        SQL = "SELECT TOP 15 Cli.IdCliente AS Codigo, Cli.CliNome AS Nome, Cli.CliCPF AS CPF, Cli.CliEmail AS Email, convert(varchar, Cli.CliDtNascimento, 103) AS DtNasc, Cli.CliCidade AS Cidade " & _
              "FROM Clientes AS Cli " & _
              "JOIN TipoClientes AS Tipo ON Cli.IdCliente = Tipo.IdCliente " & _
              "WHERE " & tipoBusca & " LIKE '%" & TxtDados.Text & "%' AND Cli.CliStatus = 0 AND Tipo.TipoFornecedor = 1 " & _
              "ORDER BY Codigo"
    End If

    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            ' SE NÃO HOUVER NADA CADASTRADO, DESABILITA O LISTVIEW
            ListViewListagemClientes.Enabled = False
        Else
            Do While rs.EOF = False
                Set Cliente = ListViewListagemClientes.ListItems.Add(, , rs("Codigo"))
                Cliente.SubItems(1) = (rs("Nome"))
                Cliente.SubItems(2) = (rs("CPF"))
                Cliente.SubItems(3) = (rs("Email"))
                Cliente.SubItems(4) = (rs("DtNasc"))
                Cliente.SubItems(5) = (rs("Cidade"))
                'SE MOVE PARA O PROXIMO REGISTRO
                rs.MoveNext
            Loop
        End If
    rs.Close

End Function

Private Sub TxtDados_KeyPress(KeyAscii As Integer)
    
    'BLOQUEIA TODOS OS CARACTERES ESPECIAIS
    If (KeyAscii >= 33 And KeyAscii <= 39) Or (KeyAscii >= 40 And KeyAscii <= 44) Or (KeyAscii >= 46 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 94) Or (KeyAscii >= 123 And KeyAscii <= 126) Or (KeyAscii >= 162 And KeyAscii <= 163) Or KeyAscii = 95 Or KeyAscii = 168 Or KeyAscii = 172 Or KeyAscii = 176 Or (KeyAscii >= 178 And KeyAscii <= 180) Or KeyAscii = 185 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
End Sub

Private Function finalizaForm()
    
    ' FINALIZA O FORM
    Unload Me
    
End Function
