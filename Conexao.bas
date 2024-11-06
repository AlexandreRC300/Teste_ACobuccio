Attribute VB_Name = "Conexao"
Public ConData As New ADODB.Connection
Sub Conectar()

    ConData.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=bdTeste;Data Source=.\SQLEXPRESS"
'Server=localhost\SQLEXPRESS;Database=master;Trusted_Connection=True;
End Sub
Sub main()
   
    Frmmain.TxtNomeServidor.Text = ".\SQLEXPRESS" 'servidor local sql server
    Frmmain.Show
End Sub
Public Sub gerarLog(sMensagem As String)
Dim registros As Integer

On Error GoTo ErrMetodo

Open "c:\temp\log" & Format(Now(), "DDMMYYYYhhmmss") & ".txt" For Output As #1

Print #1, sMensagem

Close #1

    Exit Sub

ErrMetodo:
    
    Call gerarLog(Err.Number & " - " & Err.Description & " - cmdInserir_Click")
    Exit Sub

End Sub
