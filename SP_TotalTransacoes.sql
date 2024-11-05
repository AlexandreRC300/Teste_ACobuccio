CREATE OR ALTER PROCEDURE SP_TotalTransacoes
@Data_Inicial date, 
@Data_Final date, 
@Status_Transacao char(20)
AS
	SELECT Numero_Cartao, sum(valor_transacao) as valor_total, count(id_transacao) as Quantidade_Transacoes
	FROM tb_transacoes 
	WHERE 1=1
	AND Data_Transacao >= @Data_Inicial 
	AND Data_Transacao <= @Data_Final 
	AND Status_Transacao =@Status_Transacao 
Group by Numero_Cartao 
GO
