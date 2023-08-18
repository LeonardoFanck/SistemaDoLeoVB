VERSION 5.00
Begin VB.Form FormCadastroFormaPGTO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro Forma de Pagamento"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton BtnAvancaItem 
      Caption         =   ">"
      Height          =   555
      Left            =   5505
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   135
      Width           =   675
   End
   Begin VB.CommandButton BtnVoltaItem 
      Caption         =   "<"
      Height          =   555
      Left            =   4575
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   135
      Width           =   675
   End
   Begin VB.CommandButton BtnPesquisarFormaPGTO 
      Caption         =   "->"
      Height          =   450
      Left            =   1770
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   225
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
      Left            =   7200
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   165
      Width           =   1530
   End
   Begin VB.TextBox TxtNomeFormaPGTO 
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
      Left            =   1785
      TabIndex        =   6
      Top             =   1065
      Width           =   8310
   End
   Begin VB.TextBox TxtIdFormaPGTO 
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
      Left            =   2640
      TabIndex        =   5
      Top             =   150
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      Height          =   1230
      Left            =   435
      TabIndex        =   0
      Top             =   1995
      Width           =   9645
      Begin VB.CommandButton BtnNovo 
         Caption         =   "Novo"
         Height          =   660
         Left            =   5160
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   330
         Width           =   1755
      End
      Begin VB.CommandButton BtnCancelar 
         Caption         =   "Cancelar"
         Height          =   660
         Left            =   2760
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   330
         Width           =   1755
      End
      Begin VB.CommandButton BtnConfirmar 
         Caption         =   "Confirmar"
         Height          =   660
         Left            =   375
         TabIndex        =   2
         Top             =   330
         Width           =   1755
      End
      Begin VB.CommandButton BtnAlterar 
         Caption         =   "Alterar"
         Height          =   660
         Left            =   7500
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   345
         Width           =   1755
      End
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
      Left            =   -15
      TabIndex        =   12
      Top             =   1080
      Width           =   1605
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
      Left            =   -15
      TabIndex        =   11
      Top             =   240
      Width           =   1605
   End
End
Attribute VB_Name = "FormCadastroFormaPGTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public STATUS As Integer

Private Sub BtnAlterar_Click()

    ' ALTERA PARA O STATUS 2 -> ALTERAÇÃO DE REGISTRO JÁ CADASTRADA
    STATUS = 2
    ' CHAMA A FUNÇÃO QUE PREENCHE OS DADOS
    preencheFormaPGTO (TxtIdFormaPGTO.Text)
    
End Sub

Private Sub BtnAlterar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnAvancaItem_Click()

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim ID As Integer
    
    SQL = "DECLARE @ID INT " & _
          "SET @ID = " & TxtIdFormaPGTO.Text & " " & _
          "SELECT TOP 1 IdFormaPGT FROM FormaPGTO WHERE IdFormaPGT > @ID ORDER BY IdFormaPGT"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            ID = rs("IdFormaPGT")
        Else
            ID = TxtIdFormaPGTO.Text
        End If
    rs.Close
    
    TxtIdFormaPGTO.Text = ID
    preencheFormaPGTO (ID)

End Sub

Private Sub BtnAvancaItem_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnCancelar_Click()

    On Error GoTo TrataErro
    
    ' CHAMA O FUNÇÃO QUE FINALIZA O FORM
    finalizaForm
    
Exit Sub
TrataErro:
    MsgBox " Ocorreu um erro ao tentar cancelar o cadastro N°: " & TxtIdFormaPGTO.Text & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description
End Sub

Private Sub BtnCancelar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnConfirmar_Click()

    On Error GoTo TrataError

    Dim rs As New ADODB.Recordset
    Dim statusFormaPGTO As Integer
    Dim Retorno As Integer
    Dim ID As Integer
    Dim validacao As Integer
    
    ' VERIFICA SE O NOME JÁ EXITE CADASTRADO
    SQL = "SELECT NomeFormaPGT " & _
          "FROM FormaPGTO " & _
          "WHERE NomeFormaPGT LIKE '" & TxtNomeFormaPGTO.Text & "' AND IdFormaPGT != " & TxtIdFormaPGTO.Text
    
    rs.Open SQL, cn, adOpenStatic
        ' NÃO EXISTE NENHUM NOME IGUAL CADASTRADO
        If rs.EOF = True Then
            validacao = 0
        ' EXISTE UM NOME IGUAL CADASTRADO
        Else
            validacao = 1
        End If
    rs.Close
    
    If TxtNomeFormaPGTO.Text = "" Then
        MsgBox ("Necessário informar o nome da forma de Pagamento!")
        TxtNomeFormaPGTO.SetFocus
    ElseIf validacao = 1 Then
        MsgBox ("Nome já cadastrado, informe um novo nome!")
        TxtNomeFormaPGTO.SetFocus
    Else
        If ChkInativo.Value = Checked Then
            statusFormaPGTO = 1
        ElseIf ChkInativo.Value = Unchecked Then
            statusFormaPGTO = 0
        End If
        
        If STATUS = 0 Then
            Retorno = CadastroFormaPGTO(-1, TxtNomeFormaPGTO.Text, statusFormaPGTO)
        ElseIf STATUS = 2 Then
            Retorno = CadastroFormaPGTO(TxtIdFormaPGTO.Text, TxtNomeFormaPGTO.Text, statusFormaPGTO)
        End If
        
        If Retorno = 0 Then
            MsgBox ("Forma de Pagamento atualizada com sucesso!")
        ElseIf Retorno = 1 Then
            MsgBox ("Forma de Pagamento cadastrada com sucesso!")
        ElseIf Retorno = 2 Then
            If STATUS = 0 Then
                MsgBox ("Ocorreu um erro ao tentar cadastrar a Forma de Pagamento!")
                Exit Sub
            ElseIf STATUS = 2 Then
                MsgBox ("Ocorreu um erro ao tentar atualizar a Forma de Pagamento!")
                Exit Sub
            End If
        End If
        
        If STATUS = 0 Then
            SQL = "DECLARE @MaxID INT " & _
                  "SELECT @MaxID = MAX(IdFormaPGT) FROM FormaPGTO " & _
                  "SELECT @MaxID AS FormaPGTO"
        
            rs.Open SQL, cn, adOpenStatic
                If IsNull(rs("FormaPGTO")) Then
                    STATUS = 0
                    ID = 1
                    TxtIdFormaPGTO.Text = 1
                Else
                    ' SETO O STATUS DA TELA PARA 1 -> REGISTRO JÁ CADASTRADO
                    STATUS = 1
                    ID = rs("FormaPGTO")
                End If
            rs.Close
            
            ' LIMPA O RECORDSET
            Set rs = Nothing
        ElseIf STATUS = 2 Then
            STATUS = 1
            ID = TxtIdFormaPGTO.Text
        End If
        
        ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
        preencheFormaPGTO (ID)
    End If
    
Exit Sub
TrataError:
    MsgBox "Ocorreu um erro ao finalizar o cadastro: " & TxtIdFormaPGTO.Text & vbCrLf & _
           "Erro número : " & Err.Number & vbCrLf & _
           "Detalhes : " & Err.Description
End Sub

Private Sub BtnConfirmar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Public Sub BtnNovo_Click()

    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim ID As Integer
    
    SQL = "SELECT MAX(IdFormaPGT)+1 AS FormaPGTO FROM FormaPGTO"
    
    rs.Open SQL, cn, adOpenStatic
        If IsNull(rs("FormaPGTO")) = True Then
            ID = 1
            TxtIdFormaPGTO.Text = 1
        Else
            ID = rs("FormaPGTO")
            TxtIdFormaPGTO.Text = rs("FormaPGTO")
        End If
    rs.Close
    
    ' DEFINO O STATUS DA TELA PARA 0 -> REGISTRO NOVO
    STATUS = 0
    preencheFormaPGTO (ID)

End Sub

Private Sub BtnNovo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnPesquisarFormaPGTO_Click()

    MDIFormInicio.Enabled = False
    FormBuscaFormaPGTO.FORMULARIO = "FormCadastroFormaPGTO"
    FormBuscaFormaPGTO.Show

End Sub

Private Sub BtnPesquisarFormaPGTO_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub BtnVoltaItem_Click()

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim ID As Integer
    
    SQL = "DECLARE @ID INT " & _
          "SET @ID = " & TxtIdFormaPGTO.Text & " " & _
          "SELECT TOP 1 IdFormaPGT FROM FormaPGTO WHERE IdFormaPGT < @ID ORDER BY IdFormaPGT DESC"
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            ID = rs("IdFormaPGT")
        Else
            ID = TxtIdFormaPGTO.Text
        End If
    rs.Close
    
    TxtIdFormaPGTO.Text = ID
    preencheFormaPGTO (ID)

End Sub

Private Sub BtnVoltaItem_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub ChkInativo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub Form_Load()

    On Error GoTo TrataErro
    
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim ID As Integer
    
    SQL = "DECLARE @MaxID INT " & _
          "SELECT @MaxID = MAX(IdFormaPGT) FROM FormaPGTO " & _
          "SELECT @MaxID AS FormaPGTO"
    
    rs.Open SQL, cn, adOpenStatic
        If IsNull(rs("FormaPGTO")) Then
            STATUS = 0
            ID = 1
            TxtIdFormaPGTO.Text = 1
        Else
            ' SETO O STATUS DA TELA PARA 1 -> CATEGORIA JÁ CADASTRADA
            STATUS = 1
            ID = rs("FormaPGTO")
        End If
    rs.Close
    
    ' LIMPA O RECORDSET
    Set rs = Nothing
    
    ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
    preencheFormaPGTO (ID)
    
Exit Sub
TrataErro:
    MsgBox "Algum erro ocorreu ao carregar o Form - " & Me.Name & vbCrLf & _
    " Erro número : " & Err.Number & vbCrLf & _
    " Detalhes : " & Err.Description
End Sub

Private Function preencheFormaPGTO(ID As Integer)

    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    ' NOVA CATEGORIA
    If STATUS = 0 Then
        
        ' LIBERAEÇÕES E BLOQUEIO DE CAMPOS
        TxtIdFormaPGTO.Enabled = False
        BtnAvancaItem.Enabled = False
        BtnVoltaItem.Enabled = False
        TxtNomeFormaPGTO.Enabled = True
        If FormLogin.ADMIN = 1 Then
            ChkInativo.Enabled = True
        ElseIf FormLogin.ADMIN = 2 Then
            ChkInativo.Enabled = False
        End If
        BtnNovo.Enabled = False
        BtnCancelar.Enabled = True
        BtnConfirmar.Enabled = True
        BtnAlterar.Enabled = False
        
        ' LIMPAR CAMPOS
        TxtNomeFormaPGTO.Text = ""
        ChkInativo.Value = Unchecked
        
        ' DEFINO O FOCO PARA O TXT NOME
        SendKeys ("{tab}")
        
    ' CATEGORIA JÁ CADASTRADA
    ElseIf STATUS = 1 Then
    
        SQL = "SELECT * FROM FormaPGTO " & _
              "WHERE IdFormaPGT = " & ID

        rs.Open SQL, cn, adOpenStatic
            TxtIdFormaPGTO.Text = rs("IdFormaPGT")
            TxtNomeFormaPGTO.Text = rs("NomeFormaPGT")
            If rs("StatusFormaPGT") = False Then
                ChkInativo.Value = Unchecked
            ElseIf rs("StatusFormaPGT") = True Then
                ChkInativo.Value = Checked
            End If
        rs.Close
        Set rs = Nothing
        
        STATUS = 1
        
        ' LIBERAEÇÕES E BLOQUEIO DE CAMPOS
        TxtIdFormaPGTO.Enabled = True
        BtnAvancaItem.Enabled = True
        BtnVoltaItem.Enabled = True
        TxtNomeFormaPGTO.Enabled = False
        ChkInativo.Enabled = False
        BtnNovo.Enabled = True
        BtnCancelar.Enabled = False
        BtnConfirmar.Enabled = False
        BtnAlterar.Enabled = True
    
    ElseIf STATUS = 2 Then
        ' LIBERAEÇÕES E BLOQUEIO DE CAMPOS
        TxtIdFormaPGTO.Enabled = False
        BtnAvancaItem.Enabled = False
        BtnVoltaItem.Enabled = False
        TxtNomeFormaPGTO.Enabled = True
        If FormLogin.ADMIN = 1 Then
            ChkInativo.Enabled = True
        ElseIf FormLogin.ADMIN = 2 Then
            ChkInativo.Enabled = False
        End If
        BtnNovo.Enabled = False
        BtnCancelar.Enabled = True
        BtnConfirmar.Enabled = True
        BtnAlterar.Enabled = False
    End If
    
End Function

Private Function finalizaForm()

    Dim rs As New ADODB.Recordset
    Dim ID As Integer
    Dim SQL As String
    Dim validacao As Integer
    
    ' NOVO CADASTRO
    If STATUS = 0 Then
        validacao = MsgBox("Deseja cancelar o cadastro? Todos os dados serão perdidos!", vbYesNo)
        
        If validacao = vbYes Then
            ' VOLTA O STATUS PARA CATEGORIA JÁ CADASTRADA
            STATUS = 1
            
            SQL = "DECLARE @MaxID INT " & _
                  "SELECT @MaxID = MAX(IdFormaPGT) FROM FormaPGTO " & _
                  "SELECT @MaxID AS FormaPGTO"
            
            rs.Open SQL, cn, adOpenStatic
                If IsNull(rs("FormaPGTO")) Then
                    STATUS = 0
                    ID = 1
                    TxtIdFormaPGTO.Text = 1
                Else
                    ' SETO O STATUS DA TELA PARA 1 -> CATEGORIA JÁ CADASTRADA
                    STATUS = 1
                    ID = rs("FormaPGTO")
                End If
            rs.Close
            
            ' LIMPA O RECORDSET
            Set rs = Nothing
            
            ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
            preencheFormaPGTO (ID)
            
        Else
            ' NÃO CANCELOU O PEDIDO
            TxtNomeFormaPGTO.SetFocus
        End If
    
    ' CATEGORIA JÁ CADATRADA
    ElseIf STATUS = 2 Then
        validacao = MsgBox("Deseja cancelar a alteração? Todos as mudanças serão perdidas!", vbYesNo)
        
        If validacao = vbYes Then
            ' VOLTA O STATUS PARA CATEGORIA JÁ CADASTRADA
            STATUS = 1
            
            ID = TxtIdFormaPGTO.Text
            
            ' CHAMA A FUNÇÃO QUE PREENCHE OS CAMPOS
            preencheFormaPGTO (ID)
            
        Else
            ' NÃO CANCELOU O PEDIDO
            TxtIdFormaPGTO.SetFocus
        End If
    Else
        Unload Me
    End If
End Function

Private Function CadastroFormaPGTO(IdFormaPGTO As Integer, nomeFormaPGTO As String, statusFormaPGTO As Integer)

    Dim CMD As New ADODB.Command
    Dim validacao As Integer
    
    ' CONFIGURA O COMMAND
    CMD.ActiveConnection = cn
    CMD.CommandText = "CadastroFormaPGTO"
    CMD.CommandType = adCmdStoredProc
    
    ' PARAMETROS PARA A SP
    CMD.Parameters.Append CMD.CreateParameter("RetornoOperacao", adSmallInt, adParamReturnValue, , 99)
    CMD.Parameters.Append CMD.CreateParameter("OUTPUT", adSmallInt, adParamOutput, , 99)
    CMD.Parameters.Append CMD.CreateParameter("IdFormaPGTO", adInteger, adParamInput, 12, IdFormaPGTO)
    CMD.Parameters.Append CMD.CreateParameter("NomeFormaPGTO", adVarChar, adParamInput, 50, nomeFormaPGTO)
    CMD.Parameters.Append CMD.CreateParameter("StatusFormaPGTO", adSmallInt, adParamInput, 1, statusFormaPGTO)
    
    CMD.Execute
    
    validacao = CMD.Parameters("RetornoOperacao").Value
    
    CadastroFormaPGTO = validacao
End Function

Private Sub TxtIdFormaPGTO_GotFocus()

    With TxtIdFormaPGTO
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TxtIdFormaPGTO_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 115 Then ' F4
        BtnPesquisarFormaPGTO_Click
    ElseIf KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub

Private Sub TxtIdFormaPGTO_KeyPress(KeyAscii As Integer)

    ' Verifica se o código da tecla pressionada não é um número (0-9), ENTER e BACKSPACE
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 13) Then
        ' Ignora a tecla pressionada definindo o valor de KeyAscii para 0
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        If TxtIdFormaPGTO.Text = "" Then
            MsgBox ("Necessário informar um ID!")
            KeyAscii = 0
        Else
            If STATUS = 1 Then
                ' FUNÇÃO QUE VERIFICA SE EXISTE ESTE PEDIDO NO SQL
                buscaFormaPGTO (TxtIdFormaPGTO.Text)
                KeyAscii = 0
            End If
        End If
    End If

End Sub

Public Function buscaFormaPGTO(ID As Integer)

    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim rsDados As New ADODB.Recordset
    
    SQL = "SELECT IdFormaPGT " & _
          "FROM FormaPGTO " & _
          "WHERE IdFormaPGT = " & ID
    
    rs.Open SQL, cn, adOpenStatic
        If rs.EOF = False Then
            
        Else
            MsgBox "Forma de pagamento " & ID & " não encontrada"
            
            SQL = "DECLARE @maxID INT " & _
                  "SELECT @maxID = MAX(IdFormaPGT) FROM FormaPGTO " & _
                  "SELECT @maxID AS FormaPGTO"
          
            rsDados.Open SQL, cn, adOpenStatic
                ID = rsDados("FormaPGTO")
            rsDados.Close
        End If
    rs.Close
    
    STATUS = 1
    preencheFormaPGTO (ID)
    
    SendKeys "{tab}"   ' Set the focus to the next control.
    ' VOLTA O FOCO PARA O TXT ID
    SendKeys "+{tab}" ' SHIFT TAB

End Function

Private Sub TxtNomeFormaPGTO_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then ' ESC
        finalizaForm
    End If

End Sub


Private Sub TxtNomeFormaPGTO_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then ' ENTER
        SendKeys ("{TAB}")
        KeyAscii = 0
    End If

End Sub
