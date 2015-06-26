if OBJECT_ID('dbo.Modelos') IS NOT NULL
	Drop Table dbo.Modelos
Go

Create Table dbo.Modelos (
	Empresa Char(2) Not Null,
	Codigo Numeric(10) Not Null,
	Descricao Varchar(40) Not Null,
	Abreviatura Varchar(10) Not Null,
	Ativo Char(1) Not Null
)
Alter Table dbo.Modelos Add Constraint PK_MODELOS Primary Key (Empresa, Codigo)
Alter Table dbo.Modelos Add Constraint CK_MODELOS_ATIVO Check (Ativo IN ('S', 'N'))
Go