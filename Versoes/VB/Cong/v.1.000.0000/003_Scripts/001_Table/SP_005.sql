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

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20101, 'Cadastro de Marcas', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20101, 'Cadastro de Marcas', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20101, 'Cadastro de Marcas', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20101, 'Cadastro de Marcas', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20101, 'Cadastro de Marcas', 6)
GO

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20102, 'Cadastro de Modelos', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20102, 'Cadastro de Modelos', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20102, 'Cadastro de Modelos', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20102, 'Cadastro de Modelos', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20102, 'Cadastro de Modelos', 6)
GO

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20103, 'Cadastro de Unidades', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20103, 'Cadastro de Unidades', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20103, 'Cadastro de Unidades', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20103, 'Cadastro de Unidades', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20103, 'Cadastro de Unidades', 6)
GO

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20104, 'Cadastro de Grupos', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20104, 'Cadastro de Grupos', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20104, 'Cadastro de Grupos', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20104, 'Cadastro de Grupos', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20104, 'Cadastro de Grupos', 6)
GO

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20105, 'Cadastro de SubGrupos', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20105, 'Cadastro de SubGrupos', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20105, 'Cadastro de SubGrupos', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20105, 'Cadastro de SubGrupos', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20105, 'Cadastro de SubGrupos', 6)
GO

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20200, 'Cadastro de Produtos', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20200, 'Cadastro de Produtos', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20200, 'Cadastro de Produtos', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20200, 'Cadastro de Produtos', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20200, 'Cadastro de Produtos', 6)
GO

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20401, 'Cadastro de Fornecedores', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20401, 'Cadastro de Fornecedores', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20401, 'Cadastro de Fornecedores', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20401, 'Cadastro de Fornecedores', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20401, 'Cadastro de Fornecedores', 6)
GO