IF object_id('dbo.ProdutosFornecedores') IS NOT NULL
	DROP TABLE dbo.ProdutosFornecedores
GO

CREATE TABLE dbo.ProdutosFornecedores (
	Empresa CHAR(2) NOT NULL,
	CodProduto NUMERIC(10) NOT NULL,
	CodFornecedor NUMERIC(10) NOT NULL, 
	ValorCompra NUMERIC(13, 2) NOT NULL,
	Lucro NUMERIC(5, 2) NOT NULL,
	ValorLucro NUMERIC(13, 2) NOT NULL
)

ALTER TABLE dbo.ProdutosFornecedores ADD CONSTRAINT PK_PRODUTOSFORNECEDORES PRIMARY KEY (Empresa, CodProduto, CodFornecedor)