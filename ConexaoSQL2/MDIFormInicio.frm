VERSION 5.00
Begin VB.MDIForm MDIFormInicio 
   BackColor       =   &H8000000F&
   Caption         =   "Inicio"
   ClientHeight    =   8085
   ClientLeft      =   6435
   ClientTop       =   3945
   ClientWidth     =   16065
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu CadastroCliente 
         Caption         =   "Cadastro Cliente"
      End
      Begin VB.Menu CadastroProduto 
         Caption         =   "Cadastro Produto"
      End
      Begin VB.Menu Categoria 
         Caption         =   "Cadastro Categoria"
      End
      Begin VB.Menu CadastroFormaPGTO 
         Caption         =   "Cadastro Forma de Pagamento"
      End
      Begin VB.Menu TabelaUsuarios 
         Caption         =   "Tabela Usuários"
      End
      Begin VB.Menu Operador 
         Caption         =   "Operador"
      End
      Begin VB.Menu Entrada 
         Caption         =   "Entrada"
      End
      Begin VB.Menu Pedidos 
         Caption         =   "Pedidos"
      End
   End
   Begin VB.Menu Relatorio 
      Caption         =   "Relatorio"
   End
   Begin VB.Menu Configurações 
      Caption         =   "Configurações"
   End
End
Attribute VB_Name = "MDIFormInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ID As Integer

Private Sub CadastroFormaPGTO_Click()
    FormCadastroFormaPGTO.Show
    FormCadastroFormaPGTO.SetFocus
End Sub

Private Sub Categoria_Click()
    FormCadastroCategoria.Show
    FormCadastroCategoria.SetFocus
End Sub

Private Sub Configurações_Click()
    
    If FormLogin.ADMIN = 1 Then
        FormConfiguracoesGerais.Show
        MDIFormInicio.Enabled = False
    Else
        MDIFormInicio.Enabled = False
        FormSolicitaAcesso.FORMULARIO = "MDIFormInicio"
        FormSolicitaAcesso.Show
    End If
End Sub

Private Sub CadastroCliente_Click()
    FormCadastroCliente.Show
    FormCadastroCliente.SetFocus
End Sub

Private Sub CadastroProduto_Click()
    FormCadastroProduto.Show
    FormCadastroProduto.SetFocus
End Sub

Private Sub Entrada_Click()
    FormEntrada.Show
    FormEntrada.SetFocus
End Sub

Private Sub MDIForm_Load()

    verificarPermissoesTelas

End Sub

Private Sub Operador_Click()
    FormCadastroOperador.Show
    FormCadastroOperador.SetFocus
End Sub

Private Sub Pedidos_Click()
    FormPedidos.Show
    FormPedidos.SetFocus
End Sub

Private Sub Relatorio_Click()
    
    Dim SQL As String
    
    ' SQL com o que vai buscas no banco
    SQL = "SELECT pe.IdPedidos AS Pedido, cli.CliNome AS Cliente, pgto.NomeFormaPgt AS FormaPGTO, pe.PedidoValor AS Valor, pe.PedidoDesconto AS Desconto, pe.PedidoValorTotal AS ValorTotal " & _
          "FROM Pedido as pe " & _
          "JOIN Clientes as cli on pe.PedidoIdCli = cli.IdCliente " & _
          "JOIN FormaPgto AS pgto on pe.PedidoIdPgto = pgto.IdFormaPgt " & _
          "ORDER BY pe.IdPedidos"
          
    RelatorioPedidos.DataControlRelatorioPedidos.ConnectionString = cn 'Seta a conexão para o DAO do relatório
    RelatorioPedidos.DataControlRelatorioPedidos.Source = SQL 'Passa a String para ser executada ao abrir o relatório

    RelatorioPedidos.Show
End Sub

Private Sub TabelaUsuarios_Click()
    FormTabelaUsuarios.Show
    FormTabelaUsuarios.SetFocus
End Sub

Public Function verificarPermissoesTelas()
    
    On Error GoTo TrataErro
    
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    
    SQL = "SELECT * " & _
          "FROM OperadorPermissaoTela " & _
          "WHERE IdOperador = " & ID
          
    rs.Open SQL, cn, adOpenStatic
        
        If rs("CadastroOperador") = True Then
            Operador.Visible = True
        Else
            Operador.Visible = False
        End If
        
        If rs("CadastroCategoria") = True Then
            Categoria.Visible = True
        Else
            Categoria.Visible = False
        End If
        
        If rs("CadastroCliente") = True Then
            CadastroCliente.Visible = True
        Else
            CadastroCliente.Visible = False
        End If
        
        If rs("CadastroProduto") = True Then
            CadastroProduto.Visible = True
        Else
            CadastroProduto.Visible = False
        End If
        
        If rs("CadastroFormaPGTO") = True Then
            CadastroFormaPGTO.Visible = True
        Else
            CadastroFormaPGTO.Visible = False
        End If
        
        If rs("TabelaUsuario") = True Then
            TabelaUsuarios.Visible = True
        Else
            TabelaUsuarios.Visible = False
        End If
        
        If rs("Entrada") = True Then
            Entrada.Visible = True
        Else
            Entrada.Visible = False
        End If
        
        If rs("Pedidos") = True Then
            Pedidos.Visible = True
        Else
            Pedidos.Visible = False
        End If
        
    rs.Close

Exit Function
TrataErro:
    Pedidos.Visible = True
End Function
