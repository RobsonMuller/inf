If Object_ID('dbo.ProdutosForn') Is Not Null 
	Drop Table dbo.ProdutosForn
Go 

Create Table dbo.ProdutosForn (
	Empresa Char(2) Not Null,
	CodFornecedor Numeric(10) Not Null,
	CodProduto Numeric(10) Not Null,
	Lucro Numeric(5, 2) Not Null,
	ValorCompra Numeric(13, 2) Not Null,
	ValorLucro Numeric(13, 2) Not Null
)
Alter Table dbo.ProdutosForn Add Constraint PK_PRODUTOSFORN Primary Key (Empresa, CodFornecedor, CodProduto)
Alter Table dbo.ProdutosForn Add Constraint FK_PRODUTOSFORN_PRODUTOS Foreign Key (CodProduto) References dbo.Produtos (Codigo)
Go