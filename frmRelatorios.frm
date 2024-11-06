VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatorios 
   Caption         =   "Relatórios"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar Relatorio"
      Height          =   1095
      Left            =   4080
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin MSComCtl2.DTPicker DTPInicio 
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   144179201
         CurrentDate     =   45602
      End
      Begin VB.OptionButton optSelecionarDatas 
         Caption         =   "Selecionar datas"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton optUltimos30dias 
         Caption         =   "Ultimos 30 dias"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPFinal 
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   144179201
         CurrentDate     =   45602
      End
      Begin VB.Label Label2 
         Caption         =   "Até"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "De"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmRelatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EApp As Object
Dim EwkB As Object
Dim EwkS As Object

Private Sub cmdGerar_Click()
Dim a, b As String
Dim i As Integer
Dim Exlc As New Chart
Dim sSql As String
Dim sultimos30Dias As String
Dim rsResultado As ADODB.Recordset
Dim lContador As Long
Dim dblValorTotal As Double
Dim sCaminho As String
On Error GoTo ErrMetodo
'
' Cria a componente da classe application
' inclui um novo arquivo e uma nova planilha
'
    Set EApp = CreateObject("excel.application")
    Set EwkB = EApp.Workbooks.Add
    Set EwkS = EwkB.Sheets(1)
    '
    ' exibe a aplicação Excel
    '
    EApp.Application.Visible = True
    
    ' Preenche a primeira e a segunda coluna
    ' com alguns valores numéricos
    '
    sultimos30Dias = Format(DateAdd("d", -30, Now()), "YYYY-MM-DD")
    
    sSql = ""
    sSql = sSql & " SELECT Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status_Transacao, "
    sSql = sSql & " (select dbo.fn_Categoria(tb_transacoes.valor_transacao)) as Categoria"
    sSql = sSql & " FROM tb_transacoes"
    sSql = sSql & " WHERE 1=1"
    If optSelecionarDatas.Value = True Then
        sSql = sSql & " AND data_transacao >='" & Format(DTPInicio.Value, "YYYY-MM-DD") & "' "
        sSql = sSql & " AND data_transacao <='" & Format(DTPFinal.Value, "YYYY-MM-DD") & "' "
    ElseIf optUltimos30dias.Value = True Then
        sSql = sSql & " AND data_transacao >='" & sultimos30Dias & "' "
        sSql = sSql & " AND data_transacao <='" & Format(Now(), "YYYY-MM-DD") & "' "
    End If
    Set rsResultado = New ADODB.Recordset
    rsResultado.Open sSql, ConData, 1, 1, 1
    
    If rsResultado.EOF = False Then
        
        lContador = 2
        'Cabeçalho
        
            EwkS.Range("A1").FormulaR1C1 = "Numero Cartão"
            EwkS.Range("B1").FormulaR1C1 = "Valor Transação"
            EwkS.Range("C1").FormulaR1C1 = "Data Transação"
            EwkS.Range("D1").FormulaR1C1 = "Descricao"
            EwkS.Range("E1").FormulaR1C1 = "Status Transacao"
            EwkS.Range("F1").FormulaR1C1 = "Categoria"
        
        
        
        Do While rsResultado.EOF = False
            
            EwkS.Range("A" & lContador).FormulaR1C1 = rsResultado.Fields("Numero_Cartao").Value
            EwkS.Range("A" & lContador).NumberFormat = "00000"
            EwkS.Range("B" & lContador).FormulaR1C1 = rsResultado.Fields("Valor_Transacao").Value
            EwkS.Range("B" & lContador).NumberFormat = "$#,##0.00"
            EwkS.Range("C" & lContador).FormulaR1C1 = rsResultado.Fields("Data_Transacao").Value
            EwkS.Range("D" & lContador).FormulaR1C1 = rsResultado.Fields("Descricao").Value
            EwkS.Range("E" & lContador).FormulaR1C1 = rsResultado.Fields("Status_Transacao").Value
            EwkS.Range("F" & lContador).FormulaR1C1 = rsResultado.Fields("Categoria").Value
            
            dblValorTotal = dblValorTotal + rsResultado.Fields("Valor_Transacao").Value
            lContador = lContador + 1
            rsResultado.MoveNext
            
            
            
        Loop
        EwkS.Range("A" & lContador) = "Valor Total"
        EwkS.Range("B" & lContador) = dblValorTotal
        EwkS.Range("B" & lContador).NumberFormat = "$#,##0.00"
        
        EwkS.Range("C" & lContador) = "Qtde Transações"
        EwkS.Range("D" & lContador) = lContador
        
    End If
    
    EwkS.Columns("A:F").EntireColumn.AutoFit
     
         
    Dim x As New Shell32.Shell
    Dim f As Shell32.Folder2
       
    Set f = x.BrowseForFolder(Me.hWnd, "[Mensagem]", 16, 17)

If Not f Is Nothing Then
    sCaminho = f.Self.Path
End If
     
    EwkB.SaveAs (sCaminho & "\relatorio.xls")
    'ThisWorkbook.Path & "\Contemplados hoje" & "\" & "Relatório.xlsm")
    'contemplados_arq.Close

    
    MsgBox "Relatorio Gerado com sucesso !!", vbOKOnly
    
    Exit Sub

ErrMetodo:

    MsgBox "Ocorreu um erro ao gerar o Relatorio ! ", vbCritical
    
    Exit Sub

End Sub

Private Sub cmdSelecionarCaminho_Click()
Dim x As New Shell32.Shell
Dim f As Shell32.Folder2
      
On Error GoTo ErrMetodo

    Set f = x.BrowseForFolder(Me.hWnd, "[Mensagem]", 16, 17)
    
    If Not f Is Nothing Then
       txtCaminho.Text = f.Self.Path
    End If
    
    Exit Sub

ErrMetodo:

    MsgBox "Ocorreu um erro ao gerar o Relatorio ! ", vbCritical
    
    Exit Sub
End Sub

Private Sub Form_Load()

On Error GoTo ErrMetodo

    optUltimos30dias.Value = True
    
    DTPInicio.Enabled = False
    DTPFinal.Enabled = False
    
    Exit Sub
    

ErrMetodo:
    Call gerarLog(Err.Number & " - " & Err.Description & " - Form_Load")
    Exit Sub

End Sub


Private Sub optSelecionarDatas_Click()
On Error GoTo ErrMetodo
    DTPInicio.Enabled = True
    DTPFinal.Enabled = True
    
    Exit Sub
    
ErrMetodo:
    Call gerarLog(Err.Number & " - " & Err.Description & " - optSelecionarDatas_Click")
    Exit Sub
    
End Sub

Private Sub optUltimos30dias_Click()
On Error GoTo ErrMetodo
    DTPInicio.Enabled = True
    DTPFinal.Enabled = True
    
    Exit Sub
    
ErrMetodo:
    Call gerarLog(Err.Number & " - " & Err.Description & " - optUltimos30dias_Click")
    Exit Sub
End Sub
