VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCadastroTransacoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Transações"
   ClientHeight    =   7320
   ClientLeft      =   240
   ClientTop       =   585
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   9585
   Begin VB.Frame Frame1 
      Caption         =   "Resultado da Pesquisa"
      Height          =   2055
      Left            =   120
      TabIndex        =   29
      Top             =   3960
      Width           =   9375
      Begin MSFlexGridLib.MSFlexGrid MSFlexPesquisa 
         Height          =   1695
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2990
         _Version        =   393216
      End
   End
   Begin VB.Frame fraPesquisa 
      Caption         =   "Filtros Para Pesquisa"
      Height          =   1095
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   9375
      Begin VB.ComboBox cboConsultaStatusTransacao 
         Height          =   315
         ItemData        =   "frmCadastroTransacoes.frx":0000
         Left            =   5400
         List            =   "frmCadastroTransacoes.frx":000D
         TabIndex        =   28
         Top             =   720
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPConsultaDataTransacao 
         Height          =   375
         Left            =   5400
         TabIndex        =   27
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   144244737
         CurrentDate     =   45601
      End
      Begin VB.TextBox txtConsultaValorTransacao 
         Height          =   285
         Left            =   1680
         TabIndex        =   26
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtConsultaNumCartao 
         Height          =   285
         Left            =   1680
         TabIndex        =   25
         Top             =   300
         Width           =   1935
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "Pesquisar"
         Height          =   735
         Left            =   8040
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Status Transação :"
         Height          =   255
         Left            =   3840
         TabIndex        =   24
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Valor da Transação :"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Data da Transação :"
         Height          =   255
         Left            =   3840
         TabIndex        =   22
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Numero do Cartao :"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame frmacoes 
      Caption         =   "Ações"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   9375
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar Insercao/Edicao"
         Enabled         =   0   'False
         Height          =   735
         Left            =   6240
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   735
         Left            =   7800
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "Gravar Insercao/Edicao"
         Enabled         =   0   'False
         Height          =   735
         Left            =   4680
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdApagar 
         Caption         =   "Apagar Registro"
         Enabled         =   0   'False
         Height          =   735
         Left            =   2880
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar Registro"
         Enabled         =   0   'False
         Height          =   735
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdInserir 
         Caption         =   "Inserir Novo Registro"
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frmcampos 
      Caption         =   "Campos"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   9375
      Begin VB.ComboBox cboTransacao 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCadastroTransacoes.frx":0030
         Left            =   1440
         List            =   "frmCadastroTransacoes.frx":003D
         TabIndex        =   13
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtDescricao 
         Enabled         =   0   'False
         Height          =   855
         Left            =   1440
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1200
         Width           =   7815
      End
      Begin MSComCtl2.DTPicker DTPTransacao 
         Height          =   255
         Left            =   7800
         TabIndex        =   11
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   144244737
         CurrentDate     =   45600
      End
      Begin VB.TextBox txtValorTransacao 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtNumeroCartao 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtID_transacao 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Status Transação :"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Descrição :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Data Transação :"
         Height          =   255
         Left            =   6480
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Valor Transação :"
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Número Cartão :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "ID :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmCadastroTransacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blInserir As Boolean
Dim blEditar As Boolean

Private Sub cmdApagar_Click()
Dim sSql As String

On Error GoTo ErrMetodo

If MsgBox("Confirma a exclusão do registro ?", vbYesNo) = vbYes Then
    sSql = ""
    sSql = sSql & " DELETE FROM [dbo].[tb_transacoes]"
    sSql = sSql & "  WHERE id_transacao='" & txtID_transacao.Text & "'"
    
    ConData.Execute sSql
    
    MsgBox "Exclusão da transação executada com sucesso!", vbOKOnly
    
    Call LimparCampos
End If

    Exit Sub

ErrMetodo:
    
    Call gerarLog(Err.Number & " - " & Err.Description & " - cmdApagar_click")
    Exit Sub


End Sub

Private Sub cmdCancelar_Click()
    
On Error GoTo ErrMetodo

    Call LimparCampos
    
    cmdCancelar.Enabled = False
    
    Exit Sub
ErrMetodo:
    
    Call gerarLog(Err.Number & " - " & Err.Description & " - cmdCancelar_Click")
    Exit Sub
    
End Sub

Private Sub cmdEditar_Click()

On Error GoTo ErrMetodo

    blEditar = True
    cmdGravar.Enabled = True
   
    txtNumeroCartao.Enabled = True
    txtValorTransacao.Enabled = True
    DTPTransacao.Enabled = True
    txtDescricao.Enabled = True
    cboTransacao.Enabled = True

    Exit Sub
ErrMetodo:
    
    Call gerarLog(Err.Number & " - " & Err.Description & " - cmdEditar_Click")
    Exit Sub


End Sub

Private Sub cmdGravar_Click()
Dim sSql As String

    
On Error GoTo ErrMetodo

    If blInserir = True Then
        sSql = ""
        sSql = sSql & " INSERT INTO [dbo].[tb_transacoes]"
        sSql = sSql & " (Numero_Cartao ,Valor_Transacao "
        sSql = sSql & " ,Data_Transacao ,Descricao "
        sSql = sSql & " ,Status_Transacao)"
        sSql = sSql & " VALUES "
        sSql = sSql & " ('" & txtNumeroCartao.Text & "',"
        sSql = sSql & " '" & Replace(txtValorTransacao.Text, ",", ".") & "',"
        sSql = sSql & " '" & DTPTransacao.Value & "',"
        sSql = sSql & " '" & txtDescricao.Text & "',"
        sSql = sSql & " '" & cboTransacao.Text & "')"
        
        ConData.Execute sSql
        
        MsgBox "Transação inserida com sucesso!!", vbOKOnly
    
    ElseIf blEditar = True Then
        sSql = ""
        sSql = sSql & " UPDATE   [dbo].[tb_transacoes] "
        sSql = sSql & " SET "
        sSql = sSql & " Numero_Cartao ='" & txtNumeroCartao.Text & "' "
        sSql = sSql & " ,Valor_Transacao ='" & Replace(txtValorTransacao.Text, ",", ".") & "' "
        sSql = sSql & " ,Data_Transacao ='" & DTPTransacao.Value & "' "
        sSql = sSql & " ,Descricao ='" & txtDescricao.Text & "' "
        sSql = sSql & " ,Status_Transacao ='" & cboTransacao.Text & "' "
        sSql = sSql & "  WHERE "
        sSql = sSql & " Id_transacao = '" & txtID_transacao.Text & "' "
        
        ConData.Execute sSql
        
        MsgBox "Transação alterada com sucesso!!", vbOKOnly
        
    End If


    Call LimparCampos


    Exit Sub
ErrMetodo:
    
    Call gerarLog(Err.Number & " - " & Err.Description & " - cmdGravar_Click")
    Exit Sub
End Sub

Private Sub cmdInserir_Click()
    
On Error GoTo ErrMetodo
    Call LimparCampos
    
    cmdApagar.Enabled = False
    cmdEditar.Enabled = False
    cmdGravar.Enabled = True
    cmdCancelar.Enabled = True
    blInserir = True
    
    
    txtNumeroCartao.Enabled = True
    txtValorTransacao.Enabled = True
    DTPTransacao.Enabled = True
    txtDescricao.Enabled = True
    cboTransacao.Enabled = True
    
    
    txtNumeroCartao.SetFocus
    
    Exit Sub
ErrMetodo:
    
    Call gerarLog(Err.Number & " - " & Err.Description & " - cmdInserir_Click")
    Exit Sub
    
End Sub
Private Sub LimparCampos()

On Error GoTo ErrMetodo
    
    txtConsultaNumCartao.Text = ""
    txtConsultaValorTransacao.Text = ""
    cboConsultaStatusTransacao.Text = ""
    
    
    txtID_transacao.Text = ""
    txtNumeroCartao.Text = ""
    txtValorTransacao.Text = ""
    DTPTransacao.Value = Format(Now(), "DD-MM-YYYY")
    txtDescricao.Text = ""
    cboTransacao.Text = ""
    
    txtID_transacao.Enabled = False
    txtNumeroCartao.Enabled = False
    txtValorTransacao.Enabled = False
    DTPTransacao.Enabled = False
    txtDescricao.Enabled = False
    cboTransacao.Enabled = False

    blInserir = False
    blEditar = False
    
    cmdGravar.Enabled = False
    cmdEditar.Enabled = False
    cmdApagar.Enabled = False
    MSFlexPesquisa.Clear
    
    
    Exit Sub

ErrMetodo:
    
    Call gerarLog(Err.Number & " - " & Err.Description & " - cmdInserir_Click")
    Exit Sub

End Sub

Private Sub cmdPesquisar_Click()
Dim sSql As String
Dim blExisteParametro As Boolean
Dim lContador As Long

On Error GoTo ErrMetodo
    Dim rsPesquisa As ADODB.Recordset
    blExisteParametro = False
    
    Set rsPesquisa = New ADODB.Recordset
    sSql = ""
    sSql = sSql & " SELECT * FROM TB_transacoes "
    sSql = sSql & " WHERE 1=1 "
    
    If Len(Trim(txtConsultaNumCartao.Text)) > 0 Then
        sSql = sSql & " AND Numero_Cartao='" & txtConsultaNumCartao.Text & "'"
        blExisteParametro = True
    End If
    
    If Len(Trim(txtConsultaValorTransacao.Text)) > 0 Then
        sSql = sSql & " AND Valor_Transacao='" & txtConsultaValorTransacao.Text & "'"
        blExisteParametro = True
    End If
    
    If Len(Trim(DTPConsultaDataTransacao.Value)) > 0 Then
        sSql = sSql & " AND Data_Transacao='" & DTPConsultaDataTransacao.Value & "'"
        blExisteParametro = True
    End If
    
    If Len(Trim(cboConsultaStatusTransacao.Text)) > 0 Then
        sSql = sSql & " AND Status_transacao='" & cboConsultaStatusTransacao.Text & "'"
        blExisteParametro = True
    End If
    
    If blExisteParametro = False Then
        MsgBox "Por favor selecione algum filtro para pesquisa ", vbOKOnly
        Exit Sub
    Else
    
        rsPesquisa.Open sSql, ConData, 1, 1, 1
   
        MSFlexPesquisa.Cols = 7
        MSFlexPesquisa.ColWidth(0) = 500
        MSFlexPesquisa.TextMatrix(0, 0) = "Linha"
        For lContador = 0 To rsPesquisa.Fields.Count - 1
          MSFlexPesquisa.ColAlignment(lContador) = vbCenter
          MSFlexPesquisa.ColWidth(lContador + 1) = 1500
          MSFlexPesquisa.TextMatrix(0, lContador + 1) = rsPesquisa.Fields(lContador).Name
        Next
        MSFlexPesquisa.Rows = rsPesquisa.RecordCount + 1
        lContador = 1
        Do While Not rsPesquisa.EOF
           MSFlexPesquisa.TextMatrix(lContador, 0) = lContador
           MSFlexPesquisa.TextMatrix(lContador, 1) = rsPesquisa(0) 'Id_transacao
           MSFlexPesquisa.TextMatrix(lContador, 2) = rsPesquisa(1) 'Numero_cartao
           MSFlexPesquisa.TextMatrix(lContador, 3) = rsPesquisa(2) 'valor_transacao
           MSFlexPesquisa.TextMatrix(lContador, 4) = rsPesquisa(3) 'data_transacao
           MSFlexPesquisa.TextMatrix(lContador, 5) = rsPesquisa(4) 'Descricao
           MSFlexPesquisa.TextMatrix(lContador, 6) = rsPesquisa(5) 'status_transacao
           lContador = lContador + 1
           rsPesquisa.MoveNext
        Loop
  
    End If
    
    Exit Sub

ErrMetodo:
    
    Call gerarLog(Err.Number & " - " & Err.Description & " - cmdPesquisar_Click")
    Exit Sub
End Sub

Private Sub cmdSair_Click()
    Unload Me
    
    
    
End Sub



Private Sub MSFlexPesquisa_Click()

On Error GoTo ErrMetodo
    txtID_transacao.Text = MSFlexPesquisa.TextMatrix(MSFlexPesquisa.RowSel, 1)
    txtNumeroCartao.Text = MSFlexPesquisa.TextMatrix(MSFlexPesquisa.RowSel, 2)
    txtValorTransacao.Text = MSFlexPesquisa.TextMatrix(MSFlexPesquisa.RowSel, 3)
    DTPTransacao.Value = MSFlexPesquisa.TextMatrix(MSFlexPesquisa.RowSel, 4)
    txtDescricao.Text = MSFlexPesquisa.TextMatrix(MSFlexPesquisa.RowSel, 5)
    cboTransacao.Text = MSFlexPesquisa.TextMatrix(MSFlexPesquisa.RowSel, 6)
    
    cmdEditar.Enabled = True
    cmdApagar.Enabled = True

    Exit Sub

ErrMetodo:
    
    Call gerarLog(Err.Number & " - " & Err.Description & " - cmdInserir_Click")
    Exit Sub

End Sub

Private Sub gerarLog(sMensagem As String)
Dim registros As Integer

On Error GoTo ErrMetodo

'Set db = opendatabase(Text1.Text)
'Set rs = db.openrecordset("Authors")

'rs.MoveLast
'rs.MoveFirst
'registros = rs.RecordCount

Open "c:\temp\log" & Format(Now(), "DDMMYYYYhhmmss") & ".txt" For Output As #1

Print #1, sMensagem

Close #1

    Exit Sub

ErrMetodo:
    
    Call gerarLog(Err.Number & " - " & Err.Description & " - cmdInserir_Click")
    Exit Sub

End Sub
