IF object_id('dbo.Empresas') IS NOT NULL
	DROP TABLE dbo.Empresas 
GO

CREATE TABLE dbo.Empresas (
	Codigo CHAR(2) NOT NULL,
	RazaoSocial VARCHAR(80),
	NomeFantasia VARCHAR(80),
	CPFCNPJ CHAR(14),
	FoneCom CHAR(11),
	RamalCom CHAR(5),
	FoneCel CHAR(11),
	FoneRes CHAR(11),
	DtFundacao DATETIME NOT NULL,
	DtMovimento DATETIME NOT NULL,
	Email VARCHAR(255),
	WebSite VARCHAR(255),
	Obs VARCHAR(255),
	Situacao CHAR(1),
	PathReport VARCHAR(500)
)

ALTER TABLE dbo.Empresas ADD CONSTRAINT PK_EMPRESAS PRIMARY KEY (Codigo)
ALTER TABLE dbo.Empresas ADD CONSTRAINT CK_EMPRESAS_SITUACAO CHECK (Situacao IN ('A', 'D'))
GO

IF object_id('dbo.EmpresasContatos') IS NOT NULL
	DROP TABLE dbo.EmpresasContatos
GO

CREATE TABLE dbo.EmpresasContatos (
	Empresa CHAR(2) NOT NULL,
	Codigo NUMERIC(10) NOT NULL,
	Nome VARCHAR(80) NOT NULL,
	Setor VARCHAR(60),
	Fone CHAR(11),
	Ramal CHAR(5),
	Email VARCHAR(255),
	Obs VARCHAR(255)
)

ALTER TABLE dbo.EmpresasContatos ADD CONSTRAINT PK_EMPRESASCONTATOS PRIMARY KEY (Empresa, Codigo)
ALTER TABLE dbo.EmpresasContatos ADD CONSTRAINT FK_EMPRESASCONTATOS_EMPRESAS FOREIGN KEY (Empresa) REFERENCES dbo.Empresas (Codigo)
GO 