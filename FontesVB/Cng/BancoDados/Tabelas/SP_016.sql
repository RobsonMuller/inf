IF object_id('dbo.FornecedoresContatos') IS NOT NULL
	DROP TABLE dbo.FornecedoresContatos
GO

CREATE TABLE dbo.FornecedoresContatos (
	Empresa CHAR(2) NOT NULL,
	Codigo NUMERIC(10) NOT NULL,
	CodFornecedor NUMERIC(10) NOT NULL,
	Nome VARCHAR(80) NOT NULL,
   	Telefone CHAR(11) NULL,
	Email VARCHAR(255) NULL
)
ALTER TABLE dbo.FornecedoresContatos ADD CONSTRAINT PK_FORNECEDORESCONTATOS PRIMARY KEY (Empresa, Codigo)
GO