VERSION 5.00
Begin VB.Form FormCadastroProduto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro Produto"
   ClientHeight    =   6495
   ClientLeft      =   5325
   ClientTop       =   4650
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   10740
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton CadastroCategoria 
      Caption         =   "Cadastro de Categoria"
      Height          =   375
      Left            =   8415
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2745
      Width           =   1800
   End
   Begin VB.Frame Frame2 
      Caption         =   "Finanças"
      Height          =   1725
      Left            =   645
      TabIndex        =   18
      Top             =   2985
      Width           =   3615
      Begin VB.TextBox TxtCusto 
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
         Left            =   1635
         TabIndex        =   5
         Top             =   1020
         Width           =   1755
      End
      Begin VB.TextBox TxtValor 
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
         Left            =   1635
         TabIndex        =   4
         Top             =   270
         Width           =   1755
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Custo"
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
         Index           =   1
         Left            =   195
         TabIndex        =   20
         Top             =   1095
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor"
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
         Index           =   0
         Left            =   450
         TabIndex        =   19
         Top             =   375
         Width           =   930
      End
   End
   Begin VB.CommandButton BtnAvancaItem 
      Caption         =   ">"
      Height          =   555
      Left            =   5715
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   360
      Width           =   675
   End
   Begin VB.CommandButton BtnVoltaItem 
      Caption         =   "<"
      Height          =   555
      Left            =   4785
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   360
      Width           =   675
   End
   Begin VB.CommandButton BtnPesquisarProduto 
      Caption         =   "->"
      Height          =   450
      Left            =   1980
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   450
      Width           =   690
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
      Left            =   7410
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   390
      Width           =   1530
   End
   Begin VB.ComboBox CBoxCategoria 
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
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2100
      Width           =   8265
   End
   Begin VB.Frame Frame1 
      Height          =   1230
      Left            =   405
      TabIndex        =   14
      Top             =   5100
      Width           =   9765
      Begin VB.CommandButton BtnAlterar 
         Caption         =   "Alterar"
         Height          =   660
         Left            =   5205
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   330
         Width           =   1875
      End
      Begin VB.CommandButton BtnNovo 
         Caption         =   "Novo"
         Height          =   660
         Left            =   390
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   330
         Width           =   1875
      End
      Begin VB.CommandButton BtnCancelar 
         Caption         =   "Cancelar"
         Height          =   660
         Left            =   7470
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   330
         Width           =   1875
      End
      Begin VB.CommandButton BtnConfirmar 
         Caption         =   "Confirmar"
         Height          =   660
         Left            =   2850
         TabIndex        =   6
         Top             =   330
         Width           =   1875
      End
   End
   Begin VB.TextBox TxtEstoque 
      Alignment       =   2  'Center
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
      Left            =   4800
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1890
   End
   Begin VB.TextBox TxtNomeProduto 
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
      Left            =   1995
      TabIndex        =   2
      Top             =   1290
      Width           =   8310
   End
   Begin VB.TextBox TxtCodigoProduto 
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
      Left            =   2850
      TabIndex        =   1
      Top             =   375
      Width           =   1410
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Estoque"
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
      Left            =   5055
      TabIndex        =   13
      Top             =   3270
      Width           =   1380
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Categoria"
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
      Left            =   200
      TabIndex        =   12
      Top             =   2135
      Width           =   1600
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
      Left            =   200
      TabIndex        =   11
      Top             =   1300
      Width           =   1600
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
      Left            =   200
      TabIndex        =   0
      Top             =   465
      Width           =   1600
   End
End
Attribute VB_Name = "FormCadastroProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public STATUS As Integer

Private Sub BtnAlterar_Click()

    STATUS = 2
    preencheProduto (TxtCodigoProduto.Text)

End Sub

Private Sub BtnAvancaItem_Click()

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim codProduto As Integer
    
    SQL = "DECLARE @ID INT " & _
          "SET @ID = " & TxtCodigoProduto.Text & " " & _
          "SELECT TOP 1 CodProduto FROM Produtos WHERE CodProduto > @ID ORDER BY CodProduto"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            codProduto = rs("CodProduto")
        Else
            codProduto = TxtCodigoProduto.Text
        End If
    rs.Close
    
    TxtCodigoProduto.Text = codProduto
    preencheProduto (codProduto)

End Sub

Private Sub BtnAvancaItem_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnConfirmar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Public Sub BtnNovo_Click()
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim numeroPedido As Integer

    SQL = "SELECT MAX(CodProduto)+1 AS Produto FROM Produtos"
    
    rs.Open SQL, cn, adOpenStatic
        If IsNull(rs("Produto")) = True Then
            numeroProduto = 1
            TxtCodigoProduto.Text = 1
        Else
            numeroProduto = rs("Produto")
            TxtCodigoProduto.Text = rs("Produto")
        End If
    rs.Close
    
    ' DEFINO O STATUS DA TELA PARA 0 -> PRODUTO NOVO
    STATUS = 0
    preencheProduto (numeroProduto)
End Sub

Private Sub BtnNovo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnPesquisarProduto_Click()
    FormBuscaProduto.FORMULARIO = "FormCadastroProduto"
    FormBuscaProduto.Show
    FormBuscaProduto.SetFocus
    MDIFormInicio.Enabled = False
End Sub

Private Sub BtnVoltaItem_Click()

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim codProduto As Integer
    
    SQL = "DECLARE @ID INT " & _
          "SET @ID = " & TxtCodigoProduto.Text & " " & _
          "SELECT TOP 1 CodProduto FROM Produtos WHERE CodProduto < @ID ORDER BY CodProduto DESC"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            codProduto = rs("CodProduto")
        Else
            codProduto = TxtCodigoProduto.Text
        End If
    rs.Close
    
    TxtCodigoProduto.Text = codProduto
    preencheProduto (codProduto)

End Sub

Private Sub BtnVoltaItem_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub CadastroCategoria_Click()

    chamaCadastroCategoria

End Sub

Private Function chamaCadastroCategoria()

    FormCadastroCategoria.Show
    FormCadastroCategoria.BtnNovo_Click
    FormCadastroCategoria.TxtNomeCategoria.SetFocus

End Function

Private Sub CBoxCategoria_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub CBoxCategoria_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys ("{tab}")
        KeyAscii = 0
    End If

End Sub

Private Sub ChkInativo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub Form_Load()
    On Error GoTo TrataErro
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim numeroProduto As Integer
    
    SQL = "DECLARE @maxProduto INT " & _
          "SELECT @maxProduto = MAX(CodProduto) FROM Produtos " & _
          "SELECT @maxProduto AS Produto"
    
    rs.Open SQL, cn, adOpenStatic
        If IsNull(rs("Produto")) Then
            STATUS = 0
            numeroProduto = 1
            TxtCodigoProduto.Text = 1
        Else
            numeroProduto = rs("Produto")
            STATUS = 1
        End If
    rs.Close
    
    SQL = "SELECT NomeCategoria " & _
          "FROM Categoria " & _
          "WHERE StatusCategoria = 0"
                    
    'ADICIONANDO ITENS NO COMBO BOX
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = True Then
            MsgBox ("Nenhuma categoria cadastrada!")
        Else
            Do While rs.EOF = False
                CBoxCategoria.AddItem (rs("NomeCategoria"))
            rs.MoveNext
            Loop
        End If
    rs.Close
    
    Set rs = Nothing
    
    ' LIMPA O RECORDSET
    Set rs = Nothing

    ' CHAMA A FUNÇÃO DE PREENCHER O PEDIDO
    preencheProduto (numeroProduto)
Exit Sub

TrataErro:
    MsgBox "Algum erro ocorreu - " & Err.Description
    
End Sub

Private Sub BtnConfirmar_Click()

    On Error GoTo TrataErro
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim numeroCategoria As Integer
    Dim ID As Integer
    
    If TxtNomeProduto.Text = "" Then
        MsgBox "Preencha o campo Nome!"
        TxtNomeProduto.SetFocus
    ElseIf CBoxCategoria.Text = "" Then
        MsgBox "Preencha o campo Categoria!"
        CBoxCategoria.SetFocus
    Else
        If TxtValor.Text = "" Then
            TxtValor.Text = 0
        End If
        
        If TxtCusto.Text = "" Then
            TxtCusto.Text = 0
        End If
    
        ' CHAMA A FUNÇÃO QUE FINALIZA (ALTERA / CADASTRA)
        finalizaCadastro
        
        SQL = "SELECT MAX(CodProduto)+1 FROM Produtos"
        rs.Open SQL, cn, adOpenStatic
            TxtCodigoProduto.Text = rs.GetString 'GetStrin pois o recordSet deve ser transformado em uma string para ir para o Text
        rs.Close
        
        Set rs = Nothing
    
        SQL = "DECLARE @maxID INT " & _
              "SELECT @maxID = MAX(CodProduto) FROM Produtos " & _
              "SELECT @maxID AS Produto"

        rs.Open SQL, cn, adOpenStatic
            If IsNull(rs("Produto")) Then
                STATUS = 0
                ID = 1
                TxtCodigoProduto.Text = 1
            Else
                ID = rs("Produto")
                STATUS = 1
            End If
        rs.Close
        
        ' LIMPA O RECORDSET
        Set rs = Nothing
    
        ' CHAMA A FUNÇÃO DE PREENCHER O PEDIDO
        preencheProduto (ID)
    End If
    
Exit Sub

TrataErro:
    MsgBox "Um erro ocorreu - " & Err.Description
    
End Sub

Private Sub BtnCancelar_Click()

    ' CHAMA A FUNÇÃO QUE FINALIZA O FORM
    finalizaForm

End Sub

Private Function preencheProduto(numeroProduto As Integer)

    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    ' SE STATUS = 0 -> NOVO CADASTRO
    If STATUS = 0 Then
        
        ' HABILITO TODOS OS CAMPOS PARA PODER CADASTRAR O NOVO PRODUTO
        TxtCodigoProduto.Enabled = False
        TxtNomeProduto.Enabled = True
        CBoxCategoria.Enabled = True
        TxtEstoque.Enabled = False
        TxtValor.Enabled = True
        TxtCusto.Enabled = True
        If FormLogin.ADMIN = 1 Then
            ChkInativo.Enabled = True
        End If
        BtnNovo.Enabled = False
        BtnCancelar.Enabled = True
        BtnConfirmar.Enabled = True
        BtnAvancaItem.Enabled = False
        BtnVoltaItem.Enabled = False
        BtnAlterar.Enabled = False
        CBoxCategoria.Clear
        
        SQL = "SELECT NomeCategoria " & _
              "FROM Categoria " & _
              "WHERE StatusCategoria = 0"
                        
        'ADICIONANDO ITENS NO COMBO BOX
        rs.Open SQL, cn, adOpenStatic
            If rs.EOF = True Then
                
            Else
                Do While rs.EOF = False
                    CBoxCategoria.AddItem (rs("NomeCategoria"))
                rs.MoveNext
                Loop
            End If
        rs.Close
        Set rs = Nothing
        
        ' LIMPA TODOS OS CAMPOS
        TxtNomeProduto.Text = ""
        TxtEstoque.Text = ""
        TxtValor.Text = 0
        TxtCusto.Text = 0
        ChkInativo.Value = Unchecked
        
        ' FOCO PARA O CAMPO DO NOME
        'SendKeys ("{TAB}")
        
    ' SE STATUS PRODUTO = 1 -> PRODUTO JÁ CADASTRADO -> MOSTRA OS DADOS
    ElseIf STATUS = 1 Then
        
        SQL = "SELECT NomeCategoria " & _
              "FROM Categoria " & _
              "WHERE StatusCategoria = 0"
                        
        'ADICIONANDO ITENS NO COMBO BOX
        rs.Open SQL, cn, adOpenStatic
            If rs.EOF = True Then
                
            Else
                Do While rs.EOF = False
                    CBoxCategoria.AddItem (rs("NomeCategoria"))
                rs.MoveNext
                Loop
            End If
        rs.Close
        Set rs = Nothing
                
        SQL = "SELECT Produtos.*, cat.NomeCategoria AS NomeCategoria " & _
              "FROM Produtos " & _
              "JOIN Categoria AS cat ON Produtos.CategoriaProduto = cat.IdCategoria " & _
              "WHERE CodProduto = " & numeroProduto
        
        rs.Open SQL, cn, adOpenStatic
            TxtCodigoProduto.Text = rs("CodProduto")
            TxtNomeProduto.Text = rs("NomeProduto")
            CBoxCategoria.Text = rs("NomeCategoria")
            TxtEstoque.Text = SPsGlobais.VerificaEstoqueProduto(numeroProduto)
            TxtValor.Text = rs("ValorProduto")
            TxtCusto.Text = rs("CustoProduto")
            If rs("StatusProduto") = True Then
                ChkInativo.Value = Checked
            Else
                ChkInativo.Value = Unchecked
            End If
        rs.Close
        
        Set rs = Nothing
        
        ' DESABILITO TODOS OS CAMPOS PARA SOMENTE VISUALIZAR
        TxtCodigoProduto.Enabled = True
        TxtNomeProduto.Enabled = False
        CBoxCategoria.Enabled = False
        TxtEstoque.Enabled = False
        TxtValor.Enabled = False
        TxtCusto.Enabled = False
        BtnNovo.Enabled = True
        BtnConfirmar.Enabled = False
        ChkInativo.Enabled = False
        BtnAvancaItem.Enabled = True
        BtnVoltaItem.Enabled = True
        BtnAlterar.Enabled = True
    ElseIf STATUS = 2 Then
    
        ' HABILITO TODOS OS CAMPOS PARA PODER ALTERAR O PRODUTO
        TxtCodigoProduto.Enabled = False
        TxtNomeProduto.Enabled = True
        CBoxCategoria.Enabled = True
        TxtEstoque.Enabled = False
        TxtValor.Enabled = True
        TxtCusto.Enabled = True
        If FormLogin.ADMIN = 1 Then
            ChkInativo.Enabled = True
        End If
        BtnNovo.Enabled = False
        BtnCancelar.Enabled = True
        BtnConfirmar.Enabled = True
        BtnAvancaItem.Enabled = False
        BtnVoltaItem.Enabled = False
        BtnAlterar.Enabled = False
    End If
End Function

Private Function finalizaForm()

    Dim rs As New ADODB.Recordset
    Dim ID As Integer
    Dim SQL As String
    Dim validacao As Integer

    If STATUS = 0 Then ' PEDIDO NOVO
        validacao = MsgBox("Caso saia do cadastro de produto, os dados serão perdidos! Deseja mesmo sair?", vbYesNo)
        
        If validacao = vbYes Then
            ' VOLTO O STATUS PARA VISUALIZAÇÃO
            STATUS = 1
            
            SQL = "DECLARE @maxID INT " & _
                  "SELECT @maxID = MAX(CodProduto) FROM Produtos " & _
                  "SELECT @maxID AS Produto"
            
            rs.Open SQL, cn, adOpenStatic
                If IsNull(rs("Produto")) Then
                    STATUS = 0
                    ID = 1
                    TxtCodigoProduto.Text = 1
                Else
                    ID = rs("Produto")
                    STATUS = 1
                End If
            rs.Close
            
            ' LIMPA O RECORDSET
            Set rs = Nothing
        
            ' CHAMA A FUNÇÃO DE PREENCHER O PEDIDO
            preencheProduto (ID)
        End If
        
    ElseIf STATUS = 1 Then
        Unload Me
    End If

End Function

Private Sub TxtCodigoProduto_GotFocus()

    With TxtCodigoProduto
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtCodigoProduto_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim validacao As Boolean

    If KeyCode = 115 Then ' F4
        BtnPesquisarProduto_Click
    ElseIf KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtCodigoProduto_KeyPress(KeyAscii As Integer)
    
    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        If TxtCodigoProduto.Text = "" Then
            MsgBox ("Necessário informar um Produto!")
            KeyAscii = 0
        Else
            If STATUS = 1 Then
                ' FUNÇÃO QUE VERIFICA SE EXISTE ESTE PEDIDO NO SQL
                buscaProduto (TxtCodigoProduto.Text)
                KeyAscii = 0
            ElseIf STATUS = 0 Then
                SendKeys ("{tab}")
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub TxtCusto_GotFocus()

    With TxtCusto
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtCusto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtCusto_KeyPress(KeyAscii As Integer)

    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER, BACKSPACE, '.' e ','
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) And (KeyAscii <> 44) And (KeyAscii <> 46) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
        KeyAscii = 0
    End If

End Sub

Private Sub TxtCusto_LostFocus()

    TxtCusto.Text = Format(TxtCusto.Text, "0.00")

End Sub

Private Sub TxtEstoque_GotFocus()

    With TxtEstoque
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtEstoque_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtEstoque_KeyPress(KeyAscii As Integer)

    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If

    If KeyAscii = 13 Then
        SendKeys ("{tab}")
        KeyAscii = 0
    End If

End Sub

Private Sub TxtNomeProduto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtNomeProduto_KeyPress(KeyAscii As Integer)

    'BLOQUEIA TODOS OS CARACTERES ESPECIAIS
    If (KeyAscii >= 33 And KeyAscii <= 39) Or (KeyAscii >= 40 And KeyAscii <= 44) Or (KeyAscii >= 46 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 94) Or (KeyAscii >= 123 And KeyAscii <= 126) Or (KeyAscii >= 162 And KeyAscii <= 163) Or KeyAscii = 95 Or KeyAscii = 168 Or KeyAscii = 172 Or KeyAscii = 176 Or (KeyAscii >= 178 And KeyAscii <= 180) Or KeyAscii = 185 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
        KeyAscii = 0
    End If

End Sub

Private Sub TxtValor_GotFocus()
    
    With TxtValor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtValor_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtValor_KeyPress(KeyAscii As Integer)

    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER, BACKSPACE, '.' e ','
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) And (KeyAscii <> 44) And (KeyAscii <> 46) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        SendKeys ("{tab}")
        KeyAscii = 0
    End If

End Sub

Public Function buscaProduto(numeroProduto As Integer)

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim rsDados As New ADODB.Recordset
    
    SQL = "SELECT * " & _
          "FROM Produtos " & _
          "WHERE CodProduto = " & numeroProduto
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            
        Else
            MsgBox "Produto " & numeroProduto & " não encontrado"
            
            SQL = "DECLARE @maxProduto INT " & _
                  "SELECT @maxProduto = MAX(CodProduto) FROM Produtos " & _
                  "SELECT @maxProduto AS Produto"
          
            rsDados.Open SQL, cn, adOpenStatic
                numeroProduto = rsDados("Produto")
            rsDados.Close
        End If
    rs.Close

    ' DEFINO O STATUS DA TELA PARA 1 -> PRODUTO JÁ CADASTRADO
    STATUS = 1
    
    ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
    preencheProduto (numeroProduto)

    SendKeys "{tab}"   ' Set the focus to the next control.
    ' VOLTA O FOCO PARA O TXT ID PRODUTO
    SendKeys "+{tab}" ' SHIFT TAB

End Function

Private Sub TxtValor_LostFocus()

    TxtValor.Text = Format(TxtValor.Text, "0.00")

End Sub

Private Function finalizaCadastro()
    
    Dim statusProduto, ID As Integer
    Dim SQL As String
    Dim CMD As New ADODB.Command
    Dim rs As New ADODB.Recordset
    Dim parameter As New ADODB.parameter
    Dim validacao As Integer
           
    If ChkInativo = Checked Then
        statusProduto = 1
    Else
        statusProduto = 0
    End If
    
    If STATUS = 0 Then
        ID = -1
    Else
        ID = TxtCodigoProduto.Text
    End If
    
    SQL = "SELECT IdCategoria FROM Categoria WHERE NomeCategoria LIKE '" & CBoxCategoria.Text & "'"
    
    rs.Open SQL, cn, adOpenStatic
        numeroCategoria = rs("IdCategoria")
    rs.Close
    Set rs = Nothing
    
    ' PASSA A CONEXÃO PARA O COMMAND
    CMD.ActiveConnection = cn
    CMD.CommandText = "cadastroProduto"
    CMD.CommandType = adCmdStoredProc
    
    ' PASSA OS PARAMETROS PARA O COMMAND
    CMD.Parameters.Append CMD.CreateParameter("RetornoOperacao", adInteger, adParamReturnValue, , 99)
    CMD.Parameters.Append CMD.CreateParameter("OUTPUT", adInteger, adParamOutput, , 99)
    CMD.Parameters.Append CMD.CreateParameter("ID", adInteger, adParamInput, , ID)
    CMD.Parameters.Append CMD.CreateParameter("Nome", adVarChar, adParamInput, 50, TxtNomeProduto.Text)
    CMD.Parameters.Append CMD.CreateParameter("Categoria", adInteger, adParamInput, 14, numeroCategoria)
    Set parameter = CMD.CreateParameter("Valor", adNumeric, adParamInput)
        parameter.Precision = 18
        parameter.NumericScale = 2
        CMD.Parameters.Append parameter
        CMD.Parameters("Valor").Value = TxtValor.Text
    Set parameter = CMD.CreateParameter("Custo", adNumeric, adParamInput)
        parameter.Precision = 18
        parameter.NumericScale = 2
        CMD.Parameters.Append parameter
        CMD.Parameters("Custo").Value = TxtCusto.Text
    CMD.Parameters.Append CMD.CreateParameter("Status", adBoolean, adParamInput, 1, statusProduto)
    
    CMD.Execute
    
    validacao = CMD.Parameters("RetornoOperacao").Value
    
    If STATUS = 0 Then
        If validacao = 1 Then
            MsgBox ("Produto cadastrado com sucesso!")
        ElseIf validacao = 2 Then
            MsgBox ("Ocorreu algum erro ao tentar cadastrar o Produto!")
            Exit Function
        End If
    ElseIf STATUS = 2 Then
        If validacao = 0 Then
            ' ALTEROU O CLIENTE COM SUCESSO
            MsgBox ("Produto alterado com sucesso!")
        ElseIf validacao = 1 Then
            MsgBox ("Ocorreu algum erro ao tentar alterar o Produto!")
            Exit Function
        End If
    End If

End Function
