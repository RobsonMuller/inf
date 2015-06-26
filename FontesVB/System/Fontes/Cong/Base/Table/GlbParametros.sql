IF OBJECT_ID('dbo.GlbParametros') IS NOT NULL
	DROP TABLE dbo.GlbParametros
GO

CREATE TABLE GlbParametros (
		Empresa Char(2) Not Null,
		PathReport Varchar(255) Not NULL
)
ALTER TABLE GlbParametros Add Constraint PK_GLBPARAMETROS Primary Key (Empresa)
GO


INSERT INTO GlbParametros (Empresa, PathReport) VALUES ('01', 'E:\VM\_Infinity\System\Cong\Relatorios')
GO