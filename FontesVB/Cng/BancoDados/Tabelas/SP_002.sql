IF object_id('dbo.Usuarios') IS NOT NULL
	DROP TABLE dbo.Usuarios 
GO

CREATE TABLE dbo.Usuarios (
	Empresa CHAR(2) NOT NULL,
	Codigo NUMERIC(10) NOT NULL,
	Nome VARCHAR(40) NOT NULL,
	Usuario VARCHAR(40) NOT NULL,
	Senha VARCHAR(40) NOT NULL,
	NivelAcesso CHAR(1) NOT NULL,
	Situacao CHAR(1) NOT NULL
)

ALTER TABLE dbo.Usuarios ADD CONSTRAINT PK_USUARIOS PRIMARY KEY (Empresa, Codigo)
ALTER TABLE dbo.Usuarios ADD CONSTRAINT FK_USUARIOS_EMPRESAS FOREIGN KEY (Empresa) REFERENCES dbo.Empresas (Codigo)
ALTER TABLE dbo.Usuarios ADD CONSTRAINT CK_USUARIOS_NIVELACESSO CHECK (NivelAcesso IN ('A', 'U')) -- A = Administrador, U = Usuario
ALTER TABLE dbo.Usuarios ADD CONSTRAINT CK_USUARIOS_SITUACAO CHECK (Situacao IN ('A', 'I', 'B')) -- A = Ativo, I = Inativo, B = Bloqueado
GO


