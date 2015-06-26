if OBJECT_ID('dbo.Unidades') IS NOT NULL	
	Drop Table dbo.Unidades
Go

Create Table dbo.Unidades (
	Empresa Char(2) Not Null,
	Codigo Numeric(10) Not Null,
	Descricao Varchar(40) Not Null,
	Abreviatura Varchar(10) Not Null,
	Ativo Char(1) Not Null
)
Alter Table dbo.Unidades Add Constraint PK_UNIDADES Primary Key (Empresa, Codigo)
Alter Table dbo.Unidades Add Constraint CK_UNIDADES_ATIVO Check (Ativo IN ('S', 'N'))
Go