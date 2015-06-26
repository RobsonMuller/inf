IF OBJECT_ID('dbo.GlbInterface') IS NOT NULL	
	DROP TABLE dbo.GlbInterface
GO

CREATE TABLE dbo.GlbInterface (
	Codigo Numeric(10) Not Null,
	Descricao Varchar(40) Not Null,
	IdButton Numeric(10) Not Null
)
Alter Table dbo.GlbInterface Add Constraint PK_GLBINTERFACE Primary Key (Codigo, IdButton)
Alter Table dbo.GlbInterface Add Constraint FK_GLBINTERFACE_GLBBUTTON Foreign Key (IdButton) References GlbButton (Codigo)
GO
