If OBJECT_ID('dbo.ClientesContatos') IS NOT NULL
	Drop Table dbo.ClientesContatos
GO

Create Table dbo.ClientesContatos (
	Empresa Char(2) Not Null,
	Codigo Numeric(10) Identity Not Null,
	CodCliente Numeric(10) Not Null,
	Nome Varchar(80) Not Null,
	Telefone Char(11) Null,
	Email Varchar(255) Null 
)
Alter Table dbo.ClientesContatos Add Constraint PK_CLIENTESCONTATOS Primary Key (Empresa, Codigo, CodCliente)
GO 