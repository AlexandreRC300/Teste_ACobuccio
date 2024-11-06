VERSION 5.00
Begin VB.MDIForm MDIGerFiscal 
   BackColor       =   &H8000000C&
   Caption         =   "Gerenciador Fiscal"
   ClientHeight    =   4740
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11415
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuTransacoes 
      Caption         =   "Transacoes"
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "Relatorios"
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "MDIGerFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
On Error GoTo ErrMetodo:

    Conectar
    
ErrMetodo:

End Sub

Private Sub mnuRelatorios_Click()
    frmRelatorios.Show
End Sub

Private Sub mnuSair_Click()
    End
End Sub

Private Sub mnuTransacoes_Click()
    frmCadastroTransacoes.Show
End Sub
