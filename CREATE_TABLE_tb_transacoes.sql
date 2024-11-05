USE [BDTeste]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[tb_transacoes](
	[Id_Transacao] [int] IDENTITY(1,1) NOT NULL,
	[Numero_Cartao] [nvarchar](20) NOT NULL,
	[Valor_Transacao] [float] NULL,
	[Data_Transacao] [date] NULL,
	[Descricao] [varchar](max) NULL,
	[Status_Transacao] [nvarchar](50) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


