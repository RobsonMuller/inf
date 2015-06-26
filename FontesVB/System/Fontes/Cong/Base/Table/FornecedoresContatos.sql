If OBJECT_ID('dbo.FornecedoresContatos') IS NOT NULL
	DROP TABLE dbo.FornecedoresContatos
Go

Create Table dbo.FornecedoresContatos (
	Empresa Char(2) Not Null,
	Codigo Numeric(10) Not Null Identity,
	CodFornecedor Numeric(10) Not Null,
	Nome Varchar(80) Not Null,
	Telefone Char(11) Null,
	Email Varchar(255) Null
)
Alter Table dbo.FornecedoresContatos Add Constraint PK_FORNECEDORESCONTATOS Primary Key (Empresa, Codigo, CodFornecedor)
Go