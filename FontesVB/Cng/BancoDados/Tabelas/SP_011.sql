IF object_id('dbo.Unidades') IS NOT NULL
	DROP TABLE dbo.Marcas 
GO

CREATE TABLE dbo.Unidades (
	Empresa CHAR(2) NOT NULL,
	Codigo NUMERIC(10) NOT NULL,
	Descricao VARCHAR(40) NOT NULL,
	Abreviatura VARCHAR(20) NOT NULL,
	Situacao CHAR(1) NOT NULL
)

ALTER TABLE dbo.Unidades ADD CONSTRAINT PK_UNIDADES PRIMARY KEY (Empresa, Codigo)
ALTER TABLE dbo.Unidades ADD CONSTRAINT CK_UNIDADES_SITUACAO CHECK (Situacao IN ('1', '2'))
GO