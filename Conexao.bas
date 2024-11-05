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
End Sub
