IF object_id('dbo.GLBPermissoes') IS NOT NULL	
	DROP TABLE dbo.GLBPermissoes
GO

CREATE TABLE dbo.GLBPermissoes (
	Empresa CHAR(2) NOT NULL,
	IdUsuario NUMERIC(10) NOT NULL,
	IdInterface NUMERIC(10) NOT NULL,
	IdButton NUMERIC(10) NOT NULL
)
ALTER TABLE dbo.GLBPermissoes ADD CONSTRAINT PK_GLBPERMISSOES PRIMARY KEY (Empresa, IdUsuario, IdInterface, IdButton)
ALTER TABLE dbo.GLBPermissoes ADD CONSTRAINT FK_GLBPERMISSOES_EMPRESAS FOREIGN KEY (Empresa) REFERENCES Empresas (Codigo)
ALTER TABLE dbo.GLBPermissoes ADD CONSTRAINT FK_GLBPERMISSOES_GLBBUTTON FOREIGN KEY (IdButton) REFERENCES GLBButton (Codigo)
Go
