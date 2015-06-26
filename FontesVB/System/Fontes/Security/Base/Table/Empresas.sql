If OBJECT_ID('dbo.Empresas') IS NOT NULL 
	Drop Table dbo.Empresas
GO

Create Table dbo.Empresas (
	Codigo Char(2) Not Null,
	RazaoSocial Varchar(120) Not Null,
	NomeFantasia Varchar(80) Not Null,
	CPFCNPJ Char(14) Not Null,
	Endereco Varchar(100) Null,
	Numero Numeric(10) Null,
	Bairro Varchar(60) Null,
	Cidade Varchar(60) Null,
	Estado Char(2) Null,
	FoneRes Char(11) Null,
	FoneCel Char(11) Null,
	Email Varchar(250) Null,
	DtMovimento DateTime Not Null,
	CodContrato Varchar(20) Not Null,
	Ativo Char(1) Not Null
)
Alter Table dbo.Empresas Add Constraint PK_EMPRESAS Primary Key (Codigo)
Alter Table dbo.Empresas Add Constraint CK_EMPRESAS_ATIVO Check (Ativo IN ('S', 'N'))
GO

Insert Into dbo.Empresas (Codigo, RazaoSocial, NomeFantasia, CPFCNPJ, DtMovimento, CodContrato, Ativo) Values ('01', 'Infinity Sistema de Informação', 'Infinity', '02182135065', {d '2015-03-03'}, '20500101', 'S')
Go