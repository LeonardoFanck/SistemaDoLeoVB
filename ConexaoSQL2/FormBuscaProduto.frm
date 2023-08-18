VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FormBuscaProduto 
   Caption         =   "Busca Produto"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
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
      Left            =   2115
      TabIndex        =   2
      Top             =   135
      Width           =   5265
   End
   Begin VB.CommandButton BtnMostrarTodos 
      Caption         =   "Mostrar Todos"
      Height          =   510
      Left            =   7695
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   135
      Width           =   2160
   End
   Begin VB.ComboBox ComboSelecao 
      Height          =   315
      ItemData        =   "FormBuscaProduto.frx":0000
      Left            =   165
      List            =   "FormBuscaProduto.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   195
      Width           =   1800
   End
   Begin MSComctlLib.ListView ListViewListagemProduto 
      Height          =   4395
      Left            =   330
      TabIndex        =   3
      Top             =   795
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   1191
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Categoria"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Estoque"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FormBuscaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tipoBusca As String
Public FORMULARIO As String

Private Sub BtnMostrarTodos_Click()

    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    TxtDados.Text = ""
    
    If FORMULARIO = "FormPedidos" Or FORMULARIO = "FormEntrada" Then
        SQL = "SELECT prod.CodProduto AS Codigo, prod.NomeProduto AS Nome, cat.NomeCategoria AS Categoria, prod.ValorProduto AS Valor " & _
              "FROM Produtos AS prod " & _
              "JOIN Categoria AS cat ON prod.CategoriaProduto = cat.IdCategoria " & _
              "WHERE prod.statusProduto = 0 " & _
              "ORDER BY prod.CodProduto"
    ElseIf FORMULARIO = "FormCadastroProduto" Then
        SQL = "SELECT prod.CodProduto AS Codigo, prod.NomeProduto AS Nome, cat.NomeCategoria AS Categoria, prod.ValorProduto AS Valor " & _
              "FROM Produtos AS prod " & _
              "JOIN Categoria AS cat ON prod.CategoriaProduto = cat.IdCategoria " & _
              "ORDER BY prod.CodProduto"
    End If
          
    ' SEMPRE LIMPA O LISTVIEW
    ListViewListagemProduto.ListItems.Clear
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            
        Else
            Do While rs.EOF = False
                Set Produto = ListViewListagemProduto.ListItems.Add(, , rs("Codigo"))
                Produto.SubItems(1) = (rs("Nome"))
                Produto.SubItems(2) = (rs("Categoria"))
                Produto.SubItems(3) = (SPsGlobais.VerificaEstoqueProduto(rs("Codigo")))
                Produto.SubItems(4) = (rs("Valor"))
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
    ' CodProduto NomeProduto, CategoriaProduto,
    ' EstoqueProduto, ValorProduto
    If ComboSelecao.Text = "Codigo" Then
        TxtDados.Text = ""
        tipoBusca = "CodProduto"
    ElseIf ComboSelecao.Text = "Nome" Then
        TxtDados.Text = ""
        tipoBusca = "NomeProduto"
    ElseIf ComboSelecao.Text = "Categoria" Then
        TxtDados.Text = ""
        tipoBusca = "NomeCategoria"
    ElseIf ComboSelecao.Text = "Valor" Then
        TxtDados.Text = ""
        tipoBusca = "ValorProduto"
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
    ComboSelecao.AddItem ("Categoria")
    ComboSelecao.AddItem ("Valor")
    
    'VALOR PADRÃO AO ABRIR A TELA
    ComboSelecao.Text = "Nome"
    
    'SETA O VALOR PADRÃO DE BUSCA
    tipoBusca = "NomeProduto"
    
    ' EXECUTO A FUNÇÃO QUE MOSTRA OS 10 PRIMIROS PRODUTOS NO GRID
    primeirosProdutos

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' ATIVA NOVAMENTE O MDI
    MDIFormInicio.Enabled = True
    
    If FORMULARIO = "FormPedidos" Then
        FormPedidos.TxtCodItem.SetFocus
    ElseIf FORMULARIO = "FormCadastroProduto" Then
    
    End If
End Sub

Private Function primeirosProdutos()

    Dim rs As New ADODB.Recordset
    Dim SQL As String
     
    If FORMULARIO = "FormPedidos" Or FORMULARIO = "FormEntrada" Then
        ' SELECIONA OS PRIMEIROS 15 PRODUTOS QUE ESTIVEREM ATIVOS
        SQL = "SELECT TOP 15 prod.CodProduto AS Codigo, prod.NomeProduto AS Nome, cat.NomeCategoria AS Categoria, prod.ValorProduto AS Valor " & _
              "FROM Produtos AS prod " & _
              "JOIN Categoria AS cat ON prod.CategoriaProduto = cat.IdCategoria " & _
              "WHERE prod.statusProduto = 0 " & _
              "ORDER BY prod.CodProduto"
    ElseIf FORMULARIO = "FormCadastroProduto" Then
        ' SELECIONA OS PRIMEIROS 15 PRODUTOS QUE ESTIVEREM ATIVOS E INATIVOS
        SQL = "SELECT TOP 15 prod.CodProduto AS Codigo, prod.NomeProduto AS Nome, cat.NomeCategoria AS Categoria, prod.ValorProduto AS Valor " & _
              "FROM Produtos AS prod " & _
              "JOIN Categoria AS cat ON prod.CategoriaProduto = cat.IdCategoria " & _
              "ORDER BY prod.CodProduto"
    End If

    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            ' SE NÃO HOUVER NADA CADASTRADO, DESABILITA O LIST VIEW
            ListViewListagemProduto.Enabled = False
        Else
            Do While rs.EOF = False
                Set Produto = ListViewListagemProduto.ListItems.Add(, , rs("Codigo"))
                Produto.SubItems(1) = (rs("Nome"))
                Produto.SubItems(2) = (rs("Categoria"))
                Produto.SubItems(3) = (SPsGlobais.VerificaEstoqueProduto(rs("Codigo")))
                Produto.SubItems(4) = (rs("Valor"))
                'SE MOVE PARA O PROXIMO REGISTRO
                rs.MoveNext
            Loop
        End If
    rs.Close

End Function

Private Function VerificaForm(IDProduto As Integer)
    
    ' VERIFICA QUAL FORM CHAMOU A TELA
    If FORMULARIO = "FormPedidos" Then
        'If FormPedidos.STATUS = 0 Then
            FormPedidos.TxtCodItem.Text = IDProduto
            FormPedidos.preencheProduto
            finalizaForm
        'End If
    ElseIf FORMULARIO = "FormEntrada" Then
        If FormEntrada.STATUS = 0 Then
            FormEntrada.TxtCodItem.Text = IDProduto
            FormEntrada.preencheProduto
            finalizaForm
        End If
    ElseIf FORMULARIO = "FormCadastroProduto" Then
        If FormCadastroProduto.STATUS = 1 Then
            FormCadastroProduto.TxtCodigoProduto.Text = IDProduto
            ' ENVIA O ID DO PRODUTO PARA A FUNÇÃO QUE VERIFICA SE EXISTE
            FormCadastroProduto.buscaProduto (IDProduto)
            finalizaForm
        End If
    End If
End Function

Private Sub ListViewListagemProduto_DblClick()

    Dim IDProduto As Integer
    
    ' PEGA O ID DO PRODUTO DA LISTA
    IDProduto = ListViewListagemProduto.SelectedItem.Text
    
    ' CHAMA A FUNCAO VERIFICA QUAL FORM CHAMOU A TELA
    VerificaForm (IDProduto)
End Sub

Private Sub ListViewListagemProduto_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim IDProduto As Integer
    
    If KeyCode = 27 Then ' ESC
        finalizaForm
    ElseIf KeyCode = 13 Then ' ENTER
        ' PEGA O ID DO PRODUTO DA LISTA
        IDProduto = ListViewListagemProduto.SelectedItem.Text

        ' CHAMA A FUNCAO VERIFICA QUAL FORM CHAMOU A TELA
        VerificaForm (IDProduto)
    End If

End Sub

Private Sub TxtDados_Change()
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    ' VERIFICA QUAL FORM CHAMOU A TELA
    ' PARA A TELA DE PEDIDO NÃO IRÁ MOSTRAR OS PRODUTOS INATIVOS
    If FORMULARIO = "FormPedidos" Or FORMULARIO = "FormEntrada" Then
        SQL = "SELECT prod.CodProduto AS Codigo, prod.NomeProduto AS Nome, cat.NomeCategoria AS Categoria, prod.ValorProduto AS Valor " & _
              "FROM Produtos AS prod " & _
              "JOIN Categoria AS cat ON prod.CategoriaProduto = cat.IdCategoria " & _
              "WHERE " & tipoBusca & " LIKE '%" & TxtDados.Text & "%' AND prod.statusProduto = 0 " & _
              "ORDER BY " & tipoBusca
    ElseIf FORMULARIO = "FormCadastroProduto" Then
        ' PARA A TELA DE PRODUTO IRÁ MOSTRAR OS PRODUTOS INATIVOS
        SQL = "SELECT prod.CodProduto AS Codigo, prod.NomeProduto AS Nome, cat.NomeCategoria AS Categoria, prod.ValorProduto AS Valor " & _
              "FROM Produtos AS prod " & _
              "JOIN Categoria AS cat ON prod.CategoriaProduto = cat.IdCategoria " & _
              "WHERE " & tipoBusca & " LIKE '%" & TxtDados.Text & "%' " & _
              "ORDER BY " & tipoBusca
    End If
    ' SEMPRE LIMPA O LISTVIEW
    ListViewListagemProduto.ListItems.Clear
    
    If TxtDados.Text = "" Then
        'SE ESTIVER VAZIO A PESQUISA, BUSCA OS PRIMEIROS PRODUTOS DO SELECT
        primeirosProdutos
    Else
        rs.Open SQL, cn, adOpenStatic
            If rs.EOF = True Then
            
            Else
                Do While rs.EOF = False
                    Set Produto = ListViewListagemProduto.ListItems.Add(, , rs("Codigo"))
                    Produto.SubItems(1) = (rs("Nome"))
                    Produto.SubItems(2) = (rs("Categoria"))
                    Produto.SubItems(3) = (SPsGlobais.VerificaEstoqueProduto(rs("Codigo")))
                    Produto.SubItems(4) = (rs("Valor"))
                    'SE MOVE PARA O PROXIMO REGISTRO
                  rs.MoveNext
                Loop
            End If
        rs.Close
    End If
    
End Sub

Private Sub TxtDados_KeyPress(KeyAscii As Integer)
    
    'BLOQUEIA TODOS OS CARACTERES ESPECIAIS
    If (KeyAscii >= 33 And KeyAscii <= 39) Or (KeyAscii >= 40 And KeyAscii <= 44) Or (KeyAscii >= 46 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 94) Or (KeyAscii >= 123 And KeyAscii <= 126) Or (KeyAscii >= 162 And KeyAscii <= 163) Or KeyAscii = 95 Or KeyAscii = 168 Or KeyAscii = 172 Or KeyAscii = 176 Or (KeyAscii >= 178 And KeyAscii <= 180) Or KeyAscii = 185 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtDados_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If
End Sub

Private Function finalizaForm()
    
    ' FINALIZA O FORM
    Unload Me
    
End Function
