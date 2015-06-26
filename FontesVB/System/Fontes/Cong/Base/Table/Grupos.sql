if OBJECT_ID('dbo.Grupos') IS NOT NULL
	Drop Table dbo.Grupos
Go

Create Table dbo.Grupos (
	Empresa Char(2) Not Null,
	Codigo Numeric(10) Not Null,
	Descricao Varchar(40) Not Null,
	Abreviatura Varchar(10) Not Null,
	Ativo Char(1) Not Null
)
Alter Table dbo.Grupos Add Constraint PK_GRUPOS Primary Key (Empresa, Codigo)
Alter Table dbo.Grupos Add Constraint CK_GRUPOS_ATIVO Check (Ativo IN ('S', 'N'))
Go