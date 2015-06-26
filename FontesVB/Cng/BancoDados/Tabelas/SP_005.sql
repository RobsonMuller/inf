IF object_id('dbo.GLBInterface') IS NOT NULL
	DROP TABLE dbo.GLBInterface 
GO

CREATE TABLE dbo.GLBInterface (
	Codigo NUMERIC(10) NOT NULL,
	Descricao VARCHAR(40) NOT NULL,
	IdButton NUMERIC(10) NOT NULL
)
ALTER TABLE dbo.GLBInterface ADD CONSTRAINT PK_GLBINTERFACE PRIMARY KEY (Codigo, IdButton)
ALTER TABLE dbo.GLBInterface ADD CONSTRAINT FK_GLBINTERFACE_GLBBUTTON FOREIGN KEY (IdButton) REFERENCES GlbButton (Codigo)
GO

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usu�rios', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usu�rios', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usu�rios', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usu�rios', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usu�rios', 6)
GO 