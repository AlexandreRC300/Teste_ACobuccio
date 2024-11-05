CREATE or ALTER FUNCTION fn_Categoria(@Valor int) 
RETURNS Nchar(14) 

BEGIN 
	DECLARE @Retorno Nchar(14)

	IF @Valor IS NULL SET @Retorno = 'Zero' 
	IF @Valor >=2000 SET @Retorno = 'Premium'
	IF @Valor <2000 SET @Retorno = 'Alta'
	IF @Valor <1001 SET @Retorno = 'Media'
	IF @Valor <501 SET @Retorno = 'Baixo'
	RETURN @Retorno 
END
