VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FormBuscaOperador 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busca Operador"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
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
      Left            =   2100
      TabIndex        =   2
      Top             =   150
      Width           =   5265
   End
   Begin VB.CommandButton BtnMostrarTodos 
      Caption         =   "Mostrar Todos"
      Height          =   510
      Left            =   7560
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Width           =   2160
   End
   Begin VB.ComboBox ComboSelecao 
      Height          =   315
      ItemData        =   "FormBuscaOperador.frx":0000
      Left            =   210
      List            =   "FormBuscaOperador.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   210
      Width           =   1800
   End
   Begin MSComctlLib.ListView ListViewListagemOperador 
      Height          =   4395
      Left            =   300
      TabIndex        =   3
      Top             =   870
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
      NumItems        =   4
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
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FormBuscaOperador"
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
    Dim StatusSelect As Boolean
    Dim AdminSelect As Boolean
    
    TxtDados.Text = ""
    
    If FORMULARIO = "FormCadastroOperador" Then
        SQL = "SELECT IdOperador AS Codigo, NomeOperador AS Nome, adminOperador AS Tipo, statusOperador AS Status " & _
              "FROM Operador " & _
              "ORDER BY IdOperador"
    'ElseIf FORMULARIO = "FormCadastroCliente" Then
    '    SQL = "SELECT IdCliente AS Codigo, CliNome AS Nome, CliCPF AS CPF, CliEmail AS Email, convert(varchar, CliDtNascimento, 103) AS DtNasc, CliCidade AS Cidade " & _
    '          "FROM Clientes " & _
    '          "ORDER BY Codigo"
    End If
          
    ' LIMPA A LISTA SEMPRE
    ListViewListagemOperador.ListItems.Clear
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            
        Else
            Do While rs.EOF = False
                Set Operador = ListViewListagemOperador.ListItems.Add(, , rs("Codigo"))
                Operador.SubItems(1) = (rs("Nome"))
                AdminSelect = (rs("Tipo"))
                StatusSelect = (rs("Status"))
                If AdminSelect = False Then
                    Operador.SubItems(2) = "NORMAL"
                ElseIf AdminSelect = True Then
                    Operador.SubItems(2) = "ADMIN"
                End If
                If StatusSelect = False Then
                    Operador.SubItems(3) = "Ativo"
                ElseIf StatusSelect = True Then
                    Operador.SubItems(3) = "Inativo"
                End If
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
        tipoBusca = "IdOperador"
    ElseIf ComboSelecao.Text = "Nome" Then
        TxtDados.Text = ""
        tipoBusca = "nomeOperador"
    ElseIf ComboSelecao.Text = "Tipo" Then
        TxtDados.Text = ""
        tipoBusca = "adminOperador"
    ElseIf ComboSelecao.Text = "Status" Then
        TxtDados.Text = ""
        tipoBusca = "statusOperador"
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
    ComboSelecao.AddItem ("Tipo")
    ComboSelecao.AddItem ("Status")
    
    'VALOR PADRÃO AO ABRIR A TELA
    ComboSelecao.Text = "Nome"
    'SETA O VALOR PADRÃO DE BUSCA
    tipoBusca = "NomeOperador"
    
    ' EXECUTO A FUNÇÃO QUE MOSTRA OS 10 PRIMIROS CLIENTES NO GRID
    primeirosRegistros
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' ATIVA NOVAMENTE O MDI
    MDIFormInicio.Enabled = True
End Sub

Private Sub ListViewListagemOperador_DblClick()

    Dim ID As Integer
    
    ' PEGA O ID DO REGISTRO DA LISTA
    ID = ListViewListagemOperador.SelectedItem.Text
    
    ' CHAMA A FUNCAO VERIFICA QUAL FORM CHAMOU A TELA
    VerificaForm (ID)

End Sub

Private Sub ListViewListagemOperador_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim ID As Integer
    
    If KeyCode = 27 Then ' ESC
        finalizaForm
    ElseIf KeyCode = 13 Then ' ENTER
        ' PEGA O ID DO CLENTE DA LISTA
        ID = ListViewListagemOperador.SelectedItem.Text

        ' CHAMA A FUNCAO VERIFICA QUAL FORM CHAMOU A TELA
        VerificaForm (ID)
    End If
End Sub

Private Function VerificaForm(ID As Integer)
    
    ' VERIFICA QUAL FORM CHAMOU A TELA
    If FORMULARIO = "FormCadastroOperador" Then
        If FormCadastroOperador.STATUS = 1 Then
            FormCadastroOperador.TxtIdOperador.Text = ID
            FormCadastroOperador.buscaOperador (ID)
            finalizaForm
            FormCadastroOperador.SetFocus
        End If
    End If
End Function

Private Sub TxtDados_Change()
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim StatusSelect As Boolean
    Dim AdminSelect As Boolean
    
    If FORMULARIO = "FormCadastroOperador" Then
        SQL = "SELECT IdOperador AS Codigo, nomeOperador AS Nome, adminOperador as Tipo, statusOperador AS Status " & _
              "FROM Operador " & _
              "WHERE " & tipoBusca & " LIKE '%" & TxtDados.Text & "%' " & _
              "ORDER BY " & tipoBusca
    'ElseIf FORMULARIO = "FormPedidos" Then
    '    SQL = "SELECT IdCliente AS Codigo, CliNome AS Nome, CliCPF AS CPF, CliEmail AS Email, convert(varchar, CliDtNascimento, 103) AS DtNasc, CliCidade AS Cidade " & _
    '          "FROM Clientes " & _
    '          "WHERE " & tipoBusca & " LIKE '%" & TxtDados.Text & "%' AND CliStatus = 0 " & _
    '          "ORDER BY " & tipoBusca
    End If
    
    'SEMPRE LIMPA O LISTVIEW
    ListViewListagemOperador.ListItems.Clear
    
    If TxtDados.Text = "" Then
        'SE ESTIVER VAZIO A PESQUISA, BUSCA OS PRIMEIROS REGISTROS DO SELECT
        primeirosRegistros
    Else
        rs.Open SQL, cn, adOpenStatic
            If rs.EOF = True Then
                
            Else
                Do While rs.EOF = False
                    Set Operador = ListViewListagemOperador.ListItems.Add(, , rs("Codigo"))
                    Operador.SubItems(1) = (rs("Nome"))
                    AdminSelect = (rs("Tipo"))
                    StatusSelect = (rs("Status"))
                    If AdminSelect = False Then
                        Operador.SubItems(2) = "NORMAL"
                    ElseIf AdminSelect = True Then
                        Operador.SubItems(2) = "ADMIN"
                    End If
                    If StatusSelect = False Then
                        Operador.SubItems(3) = "Ativo"
                    ElseIf StatusSelect = True Then
                        Operador.SubItems(3) = "Inativo"
                    End If
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

Private Function primeirosRegistros()

    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim StatusSelect As Boolean
    Dim AdminSelect As Boolean
    
    If FORMULARIO = "FormCadastroOperador" Then
        SQL = "SELECT TOP 15 IdOperador AS Codigo, nomeOperador AS Nome, adminOperador as Tipo, statusOperador AS Status " & _
              "FROM Operador " & _
              "WHERE " & tipoBusca & " LIKE '%" & TxtDados.Text & "%' " & _
              "ORDER BY IdOperador"
    'ElseIf FORMULARIO = "FormPedidos" Then
    '    SQL = "SELECT TOP 15 IdCliente AS Codigo, CliNome AS Nome, CliCPF AS CPF, CliEmail AS Email, convert(varchar, CliDtNascimento, 103) AS DtNasc, CliCidade AS Cidade " & _
    '          "FROM Clientes " & _
    '          "WHERE " & tipoBusca & " LIKE '%" & TxtDados.Text & "%' AND CliStatus = 0 " & _
    '          "ORDER BY Codigo"
    End If

    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            ' SE NÃO HOUVER NADA CADASTRADO, DESABILITA O LISTVIEW
            ListViewListagemOperador.Enabled = False
        Else
            Do While rs.EOF = False
                Set Operador = ListViewListagemOperador.ListItems.Add(, , rs("Codigo"))
                Operador.SubItems(1) = (rs("Nome"))
                AdminSelect = (rs("Tipo"))
                StatusSelect = (rs("Status"))
                If AdminSelect = False Then
                    Operador.SubItems(2) = "NORMAL"
                ElseIf AdminSelect = True Then
                     Operador.SubItems(2) = "ADMIN"
                End If
                If StatusSelect = False Then
                    Operador.SubItems(3) = "Ativo"
                ElseIf StatusSelect = True Then
                    Operador.SubItems(3) = "Inativo"
                End If
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

