VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FormTabelaUsuarios 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuários"
   ClientHeight    =   6030
   ClientLeft      =   1770
   ClientTop       =   4830
   ClientWidth     =   18555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   18555
   Begin MSComctlLib.ListView ListViewUsuarios 
      Height          =   4515
      Left            =   315
      TabIndex        =   2
      Top             =   240
      Width           =   17895
      _ExtentX        =   31565
      _ExtentY        =   7964
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   661
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   4789
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CPF"
         Object.Width           =   2196
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Email"
         Object.Width           =   4577
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Dt. Nasc"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Estado"
         Object.Width           =   3572
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Cidade"
         Object.Width           =   2937
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Bairro"
         Object.Width           =   4048
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Endereço"
         Object.Width           =   4048
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Número"
         Object.Width           =   1535
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Status"
         Object.Width           =   1244
      EndProperty
   End
   Begin VB.CommandButton BtnSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16500
      TabIndex        =   1
      Top             =   5265
      Width           =   1725
   End
   Begin VB.CommandButton BtnIncluir 
      Caption         =   "Incluir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   465
      TabIndex        =   0
      Top             =   5205
      Width           =   1725
   End
End
Attribute VB_Name = "FormTabelaUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    On Error GoTo TrataErro
    
    Call subListaUsuarios
    
Exit Sub

TrataErro:
    MsgBox "ALGUM ERRO OCORREU - " & Err.Description

End Sub

Private Sub BtnIncluir_Click()
    
    FormCadastroCliente.Show
    
End Sub

Private Sub BtnReomover_Click()

    Dim teste As String
    
    On Error GoTo TrataErro
    
    MsgBox "teste = " & teste

    If MsgBox("Confirmar a exclusão desta linha?", vbYesNo) = vbYes Then
        
        Call subListaUsuarios
    End If
Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro ao localizar o operador: " & TxtLogin & vbCrLf & _
           "Erro número : " & Err.Number & vbCrLf & _
           "Detalhes : " & Err.Description
End Sub

Private Sub BtnSair_Click()
    Unload Me
End Sub

Private Sub subListaUsuarios()

    Dim cmd As New ADODB.Command
    Dim rs As ADODB.Recordset
    Dim usuarios As ListItem
    
    On Error GoTo TrataErro
    
    'Ativa a conexão do CMD
    cmd.ActiveConnection = cn
    'Define o comando a ser executado
    cmd.CommandText = "SELECT * FROM Clientes"
    'Executa o comando (CMD) e adiciona no RecordSet
    Set rs = cmd.Execute
    
    'Limpa a lista para não duplicar os itens
    ListViewUsuarios.ListItems.Clear
    'Executa o loop aonde adiciona os dados do RecordSet ao listView
    Do While rs.EOF = False
        Set usuarios = ListViewUsuarios.ListItems.Add(, , rs("IdCliente"))
        usuarios.SubItems(1) = rs("CliNome")
        usuarios.SubItems(2) = rs("CliCPF")
        usuarios.SubItems(3) = rs("CliEmail")
        usuarios.SubItems(4) = rs("CliDtNascimento")
        usuarios.SubItems(5) = rs("CliEstado")
        usuarios.SubItems(6) = rs("CliCidade")
        usuarios.SubItems(7) = rs("CliBairro")
        usuarios.SubItems(8) = rs("CliEndereco")
        usuarios.SubItems(9) = rs("CliNumero")
        
        'Faz uma verificação para alterar o dado da tabela para ativo ou inativo
        If rs("CliStatus") = False Then
            usuarios.SubItems(10) = "Ativo"
        ElseIf rs("CliStatus") = True Then
            usuarios.SubItems(10) = "Inativo"
        Else
            usuarios.SubItems(10) = "ERRO"
        End If
        
        'Move Para o próximo registro
        rs.MoveNext
    Loop
Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro ao localizar o operador: " & TxtLogin & vbCrLf & _
           "Erro número : " & Err.Number & vbCrLf & _
           "Detalhes : " & Err.Description
End Sub
