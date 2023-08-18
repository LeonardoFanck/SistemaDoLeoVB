VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FormBuscaPedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busca Pedidos"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   10035
   Begin MSComctlLib.ListView ListViewListagemPedidos 
      Height          =   4170
      Left            =   210
      TabIndex        =   2
      Top             =   1005
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   7355
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
         Text            =   "Pedido"
         Object.Width           =   1349
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cliente"
         Object.Width           =   5345
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Forma Pgto"
         Object.Width           =   3149
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   2593
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Desconto"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Valor Total"
         Object.Width           =   2566
      EndProperty
   End
   Begin VB.ComboBox ComboSelecao 
      Height          =   315
      ItemData        =   "FormBuscaPedido.frx":0000
      Left            =   45
      List            =   "FormBuscaPedido.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   300
      Width           =   1800
   End
   Begin VB.CommandButton BtnMostrarTodos 
      Caption         =   "Mostrar Todos"
      Height          =   510
      Left            =   7575
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   210
      Width           =   2160
   End
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
      Left            =   1935
      TabIndex        =   1
      Top             =   210
      Width           =   5265
   End
End
Attribute VB_Name = "FormBuscaPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tipoBusca As String

Private Sub BtnPesquisar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        Unload Me
    End If
End Sub

Private Sub BtnMostrarTodos_Click()
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    TxtDados.Text = ""
    
    SQL = "SELECT pedido.IdPedidos AS Pedido, cli.CliNome AS Nome, pgto.NomeFormaPgt AS PGTO, pedido.PedidoValor AS Valor, pedido.PedidoDesconto AS Desconto, pedido.PedidoValorTotal AS ValorTotal " & _
          "FROM Pedido AS pedido " & _
          "JOIN Clientes AS cli ON pedido.PedidoIdCli = cli.IdCliente " & _
          "JOIN FormaPgto AS pgto ON pedido.PedidoIdPgto = pgto.IdFormaPgt " & _
          "ORDER BY pedido.IdPedidos DESC"
    
    ' SEMPRE LIMPA A LISTAGEM
    ListViewListagemPedidos.ListItems.Clear
    
    ' ADICIONA OS DADOS NA LISTAGEM
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
           
        Else
            Do While rs.EOF = False
                Set Pedido = ListViewListagemPedidos.ListItems.Add(, , rs("Pedido"))
                Pedido.SubItems(1) = rs("Nome")
                Pedido.SubItems(2) = rs("PGTO")
                Pedido.SubItems(3) = rs("Valor")
                Pedido.SubItems(4) = rs("Desconto")
                Pedido.SubItems(5) = rs("ValorTotal")
            'MOVE PARA O PRÓXIMO OBJETO DA LISTA
            rs.MoveNext
            Loop
        End If
    rs.Close
        
End Sub

Private Sub BtnMostrarTodos_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then ' ESC
        FinalizaForm
    End If
    
End Sub

Private Sub ComboSelecao_Click()
    If ComboSelecao.Text = "Pedido" Then
        TxtDados.Text = ""
        tipoBusca = "pedido.IdPedidos"
    ElseIf ComboSelecao.Text = "Nome" Then
        TxtDados.Text = ""
        tipoBusca = "cli.CliNome"
    ElseIf ComboSelecao.Text = "Forma Pgto" Then
        TxtDados.Text = ""
        tipoBusca = "pgto.NomeFormaPgt"
    ElseIf ComboSelecao.Text = "Valor" Then
        TxtDados.Text = ""
        tipoBusca = "pedido.PedidoValor"
    ElseIf ComboSelecao.Text = "Desconto" Then
        TxtDados.Text = ""
        tipoBusca = "pedido.PedidoDesconto"
    ElseIf ComboSelecao.Text = "Valor Total" Then
        TxtDados.Text = ""
        tipoBusca = "pedido.PedidoValorTotal"
    End If
End Sub



Private Sub ComboSelecao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        FinalizaForm
    End If
End Sub

Private Sub Form_Load()
    'ADICIONA OS VALORE NO COMBOBOX
    ComboSelecao.AddItem "Pedido"
    ComboSelecao.AddItem "Nome"
    ComboSelecao.AddItem "Forma Pgto"
    ComboSelecao.AddItem "Valor"
    ComboSelecao.AddItem "Desconto"
    ComboSelecao.AddItem "Valor Total"
    
    'SETA O VALOR TEXTO PADRÃO
    ComboSelecao.Text = "Nome"
    'SETA O TIPO DE BUSCA PADRÃO
    tipoBusca = "cli.CliNome"
    
    'EXECUTA A FUNÇÃO PARA MOSTRAR OS 10 PRIMEIROS PEDIDOS
    primeirosDezItens
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' ATIVA NOVAMENTE O MDI
    MDIFormInicio.Enabled = True
End Sub

Private Sub ListViewListagemPedidos_DblClick()
     
     Dim numeroPedido As Integer
     
     ' PEGA O ID DO CLENTE DA LISTA
     numeroPedido = ListViewListagemPedidos.SelectedItem.Text
     
     ' CHAMA A FUNCAO VERIFICA QUAL FORM CHAMOU A TELA
     VerificaForm (numeroPedido)
End Sub

'DEBUG PARA VERIFICAR O TAMANHO DAS COLUNAS
'Private Sub ListViewListagemPedidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    Debug.Print ColumnHeader.Width
'End Sub

Private Sub ListViewListagemPedidos_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim numeroPedido As Integer
    
    If KeyCode = 27 Then ' ESC
        FinalizaForm
        FormPedidos.STATUS = 1
    ElseIf KeyCode = 13 Then ' ENTER
        ' PEGA O ID DO CLENTE DA LISTA
        numeroPedido = ListViewListagemPedidos.SelectedItem.Text

        ' CHAMA A FUNCAO VERIFICA QUAL FORM CHAMOU A TELA
        VerificaForm (numeroPedido)
    End If
End Sub

Private Function VerificaForm(numeroPedido As Integer)
    
    If FormPedidos.STATUS = 1 Then
        FormPedidos.buscaPedido (numeroPedido)
        FinalizaForm
     End If
End Function

Private Sub TxtDados_Change()
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    'SEMPRE LIMPA O LIST VIEW PARA NÃO DUPLICAR NO GRID
    ListViewListagemPedidos.ListItems.Clear
    
    SQL = "SELECT pedido.IdPedidos AS Pedido, cli.CliNome AS Nome, pgto.NomeFormaPgt AS PGTO, pedido.PedidoValor AS Valor, pedido.PedidoDesconto AS Desconto, pedido.PedidoValorTotal AS ValorTotal " & _
          "FROM Pedido AS pedido " & _
          "JOIN Clientes AS cli ON pedido.PedidoIdCli = cli.IdCliente " & _
          "JOIN FormaPgto AS pgto ON pedido.PedidoIdPgto = pgto.IdFormaPgt " & _
          "WHERE " & tipoBusca & " LIKE '%" & TxtDados.Text & "%' " & _
          "ORDER BY pedido.IdPedidos"
    
    
    
    If TxtDados.Text = "" Then
        'SE ESTIVER VAZIO A PESQUISA, MOSTRA OS 10 PRIMEIROS ITENS
        primeirosDezItens
    Else
        rs.Open SQL, cn, adOpenStatic
            If rs.EOF = True Then
                'MsgBox "Busca não encontrada no sistema"
                'CASO NÃO ENCONTRE A BUSCA, O SISTEMA VOLTA A LETRA DIGITADA
                'TxtDados.Text = Left(TxtDados.Text, (Len(TxtDados.Text) - 1))
                'POSICIONA O FOCO DO TECLADO NO FINAL DA FRASE
                'TxtDados.SelStart = Len(TxtDados.Text)
            Else
                Do While rs.EOF = False
                    Set Pedido = ListViewListagemPedidos.ListItems.Add(, , rs("Pedido"))
                    'Pedido.SubItems(0) = rs("Pedido")
                    Pedido.SubItems(1) = rs("Nome")
                    Pedido.SubItems(2) = rs("PGTO")
                    Pedido.SubItems(3) = rs("Valor")
                    Pedido.SubItems(4) = rs("Desconto")
                    Pedido.SubItems(5) = rs("ValorTotal")
                'MOVE PARA O PRÓXIMO OBJETO DA LISTA
                rs.MoveNext
                Loop
            End If
        rs.Close
    End If
End Sub

Private Sub TxtDados_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        FinalizaForm
    End If
End Sub

Private Sub primeirosDezItens()
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    
    SQL = "SELECT TOP 15 pedido.IdPedidos AS Pedido, cli.CliNome AS Nome, pgto.NomeFormaPgt AS PGTO, pedido.PedidoValor AS Valor, pedido.PedidoDesconto AS Desconto, pedido.PedidoValorTotal AS ValorTotal " & _
          "FROM Pedido AS pedido " & _
          "JOIN Clientes AS cli ON pedido.PedidoIdCli = cli.IdCliente " & _
          "JOIN FormaPgto AS pgto ON pedido.PedidoIdPgto = pgto.IdFormaPgt " & _
          "WHERE " & tipoBusca & " LIKE '%" & TxtDados.Text & "%' " & _
          "ORDER BY pedido.IdPedidos DESC"
          
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            ' SE NÃO HOUVER NADA CADASTRADO, DESABILITA O LISTVIEW
            ListViewListagemPedidos.Enabled = False
        Else
            Do While rs.EOF = False
                Set Pedido = ListViewListagemPedidos.ListItems.Add(, , rs("Pedido"))
                Pedido.SubItems(1) = rs("Nome")
                Pedido.SubItems(2) = rs("PGTO")
                Pedido.SubItems(3) = rs("Valor")
                Pedido.SubItems(4) = rs("Desconto")
                Pedido.SubItems(5) = rs("ValorTotal")
            'MOVE PARA O PRÓXIMO OBJETO DA LISTA
            rs.MoveNext
            Loop
        End If
    rs.Close
End Sub

Private Sub TxtDados_KeyPress(KeyAscii As Integer)
    'Debug.Print KeyAscii
    
    'BLOQUEIA TODOS OS CARACTERES ESPECIAIS
    If (KeyAscii >= 33 And KeyAscii <= 39) Or (KeyAscii >= 40 And KeyAscii <= 44) Or (KeyAscii >= 46 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 94) Or (KeyAscii >= 123 And KeyAscii <= 126) Or (KeyAscii >= 162 And KeyAscii <= 163) Or KeyAscii = 95 Or KeyAscii = 168 Or KeyAscii = 172 Or KeyAscii = 176 Or (KeyAscii >= 178 And KeyAscii <= 180) Or KeyAscii = 185 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
End Sub

Private Function FinalizaForm()
    
    ' FINALIZA O FORM
    Unload Me
    
End Function
