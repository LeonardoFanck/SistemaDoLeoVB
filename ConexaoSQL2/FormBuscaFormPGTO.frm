VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FormBuscaFormaPGTO 
   Caption         =   "Busca Forma Pagamento"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   10005
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox ComboSelecao 
      Height          =   315
      ItemData        =   "FormBuscaFormPGTO.frx":0000
      Left            =   75
      List            =   "FormBuscaFormPGTO.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   225
      Width           =   1800
   End
   Begin VB.CommandButton BtnMostrarTodos 
      Caption         =   "Mostrar Todos"
      Height          =   510
      Left            =   7605
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   165
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
      Left            =   1965
      TabIndex        =   0
      Top             =   165
      Width           =   5265
   End
   Begin MSComctlLib.ListView ListViewListagemFormaPGTO 
      Height          =   4395
      Left            =   240
      TabIndex        =   3
      Top             =   825
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   1191
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FormBuscaFormaPGTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tipoBusca As String
Public FORMULARIO As String

Private Sub BtnMostrarTodos_Click()
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "SELECT pgto.IdFormaPgt AS Codigo, pgto.NomeFormaPgt AS Nome " & _
          "FROM FormaPgto AS pgto " & _
          "WHERE pgto.StatusFormaPgt = 0 " & _
          "ORDER BY Codigo"
          
    ' LIMPA A LISTA SEMPRE
    ListViewListagemFormaPGTO.ListItems.Clear
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
                
        Else
            Do While rs.EOF = False
                Set FormaPGTO = ListViewListagemFormaPGTO.ListItems.Add(, , rs("Codigo"))
                FormaPGTO.SubItems(1) = (rs("Nome"))

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
        tipoBusca = "IdFormaPGTO"
    ElseIf ComboSelecao.Text = "Nome" Then
        TxtDados.Text = ""
        tipoBusca = "NomeFormaPgt"
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
    
    'VALOR PADRÃO AO ABRIR A TELA
    ComboSelecao.Text = "Nome"
    'SETA O VALOR PADRÃO DE BUSCA
    tipoBusca = "NomeFormaPgt"
    
    ' EXECUTO A FUNÇÃO QUE MOSTRA OS 10 PRIMIROS CLIENTES NO GRID
    primeirosFormaPGTO
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' ATIVA NOVAMENTE O MDI
    MDIFormInicio.Enabled = True
    
    If FORMULARIO = "FormPedidos" Then
        If FormPedidos.STATUS = 0 Then
            FormPedidos.TxtFormaPgto.SetFocus
        End If
    ElseIf FORMULARIO = "FormEntrada" Then
        If FormEntrada.STATUS = 0 Then
            FormEntrada.TxtFormaPgto.SetFocus
        End If
    ElseIf FORMULARIO = "FormBuscaFormaPGTO" Then
        If FormCadastroFormaPGTO.STATUS = 1 Then
            FormCadastroFormaPGTO.TxtIdFormaPGTO.SetFocus
        End If
    End If
End Sub

Private Sub ListViewListagemFormaPGTO_DblClick()
    
    Dim IDFormPGTO As Integer
    
    ' PEGA O ID DO CLENTE DA LISTA
    IDFormPGTO = ListViewListagemFormaPGTO.SelectedItem.Text
    
    ' CHAMA A FUNCAO VERIFICA QUAL FORM CHAMOU A TELA
    VerificaForm (IDFormPGTO)
End Sub

Private Sub ListViewListagemFormaPGTO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim IDCliente As Integer
    
    If KeyCode = 27 Then ' ESC
        finalizaForm
    ElseIf KeyCode = 13 Then ' ENTER
        ' PEGA O ID DO CLENTE DA LISTA
        IDCliente = ListViewListagemFormaPGTO.SelectedItem.Text

        ' CHAMA A FUNCAO VERIFICA QUAL FORM CHAMOU A TELA
        VerificaForm (IDCliente)
    End If
End Sub

Private Function VerificaForm(ID As Integer)
    
    ' VERIFICA QUAL FORM CHAMOU A TELA
    If FORMULARIO = "FormPedidos" Then
        If FormPedidos.STATUS = 0 Then
            FormPedidos.TxtFormaPgto.Text = ID
            FormPedidos.PegaNomeFormaPgto
            finalizaForm
        End If
    ElseIf FORMULARIO = "FormEntrada" Then
        If FormEntrada.STATUS = 0 Then
            FormEntrada.TxtFormaPgto.Text = ID
            FormEntrada.PegaNomeFormaPgto
            finalizaForm
        End If
    ElseIf FORMULARIO = "FormBuscaFormaPGTO" Then
        If FormCadastroFormaPGTO.STATUS = 1 Then
            FormCadastroFormaPGTO.TxtIdFormaPGTO.Text = ID
            FormCadastroFormaPGTO.buscaFormaPGTO (ID)
            finalizaForm
        End If
    End If
    
End Function

Private Sub TxtDados_Change()
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
          
    SQL = "SELECT pgto.IdFormaPgt AS Codigo, pgto.NomeFormaPgt AS Nome " & _
          "FROM FormaPgto AS pgto " & _
          "WHERE " & tipoBusca & " LIKE '%" & TxtDados.Text & "%' AND pgto.StatusFormaPgt = 0 " & _
          "ORDER BY " & tipoBusca

    
    'SEMPRE LIMPA O LISTVIEW
    ListViewListagemFormaPGTO.ListItems.Clear
    
    If TxtDados.Text = "" Then
        'SE ESTIVER VAZIO A PESQUISA, BUSCA OS 10 PRIMEIROS CLIENTE DO SELECT
        primeirosFormaPGTO
    Else
        rs.Open SQL, cn, adOpenStatic
            If rs.EOF = True Then
                
            Else
                Do While rs.EOF = False
                    Set FormPGTO = ListViewListagemFormaPGTO.ListItems.Add(, , rs("Codigo"))
                    FormPGTO.SubItems(1) = (rs("Nome"))

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

Private Function primeirosFormaPGTO()

    Dim rs As New ADODB.Recordset
    Dim SQL As String
          
    ' SELECIONA AS PRIMEIRA 15 FORMAS DE PAGAMENTO QUE ESTIVEREM ATIVAS
    SQL = "SELECT TOP 15 pgto.IdFormaPgt AS Codigo, pgto.NomeFormaPgt AS Nome " & _
          "FROM FormaPgto AS pgto " & _
          "WHERE pgto.StatusFormaPgt = 0 " & _
          "ORDER BY Codigo"

    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            ' SE NÃO HOUVER NADA CADASTRADO, DESABILITA O LISTVIEW
            ListViewListagemFormaPGTO.Enabled = False
        Else
            Do While rs.EOF = False
                Set FormaPGTO = ListViewListagemFormaPGTO.ListItems.Add(, , rs("Codigo"))
                FormaPGTO.SubItems(1) = (rs("Nome"))
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

