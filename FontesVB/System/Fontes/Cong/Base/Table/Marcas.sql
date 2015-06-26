if OBJECT_ID('dbo.Marcas') IS NOT NULL	
	Drop Table dbo.Marcas
Go

Create Table dbo.Marcas (
	Empresa Char(2) Not Null,
	Codigo Numeric(10) Not Null,
	Descricao Varchar(40) Not Null,
	Abreviatura Varchar(10) Not Null,
	Ativo Char(1) Not Null
)
Alter Table dbo.Marcas Add Constraint PK_MARCAS Primary Key (Empresa, Codigo)
Alter Table dbo.Marcas Add Constraint CK_MARCAS_ATIVO Check (Ativo IN ('S', 'N'))
Go	