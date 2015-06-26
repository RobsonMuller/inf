if OBJECT_ID('dbo.SubGrupos') IS NOT NULL	
	Drop Table dbo.SubGrupos
Go

Create Table dbo.SubGrupos (
	Empresa Char(2) Not Null,
	Codigo Numeric(10) Not Null,
	Descricao Varchar(40) Not Null,
	Abreviatura Varchar(10) Not Null,
	Ativo Char(1) Not Null
)
Alter Table dbo.SubGrupos Add Constraint PK_SUBGRUPOS Primary Key (Empresa, Codigo)
Alter Table dbo.SubGrupos Add Constraint CK_SUBGRUPOS_ATIVO Check (Ativo IN ('S', 'N'))
Go