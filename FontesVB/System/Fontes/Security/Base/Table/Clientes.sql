If OBJECT_ID('dbo.Clientes') IS NOT NULL
	Drop Table dbo.Clientes
Go 

Create Table dbo.Clientes (
	Codigo Numeric(10) Not Null,
	Empresa Char(2) Not Null,
	Nome Varchar(80) Not Null,
	DataCad DateTime Not Null,
	CPFCNPJ Char(14) Not Null,
	RGIE Char(12) Null,
	CEP Char(9) Null,
	TpPessoa Char(1) Not Null,
	Situacao Char(1) Not Null,
	Endereco Varchar(80) Not Null,
	Numero Numeric(5) Null,
	CodCidade Numeric(10) Not Null,
	Bairro Varchar(60) Not Null,
	Estado Char(2) Not Null,
	Telefone Char(11) Null,
	Fax Char(11) Null,
	Cel Char(11) Null,
	Email Varchar(255) Null,
	Site Varchar(255) Null
)
Alter Table dbo.Clientes Add Constraint PK_CLIENTES Primary Key (Codigo, Empresa )
Go