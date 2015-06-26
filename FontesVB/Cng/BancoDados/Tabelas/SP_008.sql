IF object_id('dbo.UsuariosExcluidos') IS NOT NULL
	DROP TABLE dbo.UsuariosExcluidos
GO

CREATE TABLE dbo.UsuariosExcluidos ( 
	Empresa CHAR (2) NOT NULL,
	Codigo NUMERIC(10) NOT NULL,
	Nome VARCHAR(40) NOT NULL,
	Usuario VARCHAR(40) NOT NULL,
	Situacao CHAR(1) NOT NULL,
	NivelAcesso CHAR(1) NOT NULL,
	DtHrExclusao DATETIME NOT NULL,
	CodUsuarioExclusao NUMERIC(10) NOT NULL
)
GO