If OBJECT_ID('dbo.ProdutosPathImg') IS NOT NULL
	Drop Table dbo.ProdutosPathImg
Go

Create Table dbo.ProdutosPathImg (
	Empresa Char(2) Not Null,
	CodProduto Numeric(10) Not Null,
	PathImg Varchar(255) Not Null
)
Alter Table dbo.ProdutosPathImg Add Constraint PK_PRODUTOSPATHIMG Primary Key (Empresa, CodProduto)
Alter Table dbo.ProdutosPathImg Add Constraint FK_PRODUTOSPATHIMG_PRODUTOS Foreign Key (Empresa, CodProduto) References dbo.Produtos (Empresa, Codigo)
Go