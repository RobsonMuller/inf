IF object_id('dbo.Modelos') IS NOT NULL
	DROP TABLE dbo.Marcas 
GO

CREATE TABLE dbo.Modelos (
	Empresa CHAR(2) NOT NULL,
	Codigo NUMERIC(10) NOT NULL,
	Descricao VARCHAR(40) NOT NULL,
	Abreviatura VARCHAR(20) NOT NULL,
	Situacao CHAR(1) NOT NULL
)

ALTER TABLE dbo.Modelos ADD CONSTRAINT PK_MODELOS PRIMARY KEY (Empresa, Codigo)
ALTER TABLE dbo.Modelos ADD CONSTRAINT CK_MODELOS_SITUACAO CHECK (Situacao IN ('1', '2'))
GO