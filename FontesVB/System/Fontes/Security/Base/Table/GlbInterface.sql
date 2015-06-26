if OBJECT_ID('dbo.GlbInterface') IS NOT NULL	
	Drop Table dbo.GlbInterface
Go

Create Table dbo.GlbInterface (
	Codigo Numeric(10) Not Null,
	Descricao Varchar(40) Not Null,
	IdButton Numeric(10) Not Null
)
Alter Table dbo.GlbInterface Add Constraint PK_GLBINTERFACE Primary Key (Codigo, IdButton)
Alter Table dbo.GlbInterface Add Constraint FK_GLBINTERFACE_GLBBUTTON Foreign Key (IdButton) References GlbButton (Codigo)
Go

-- Usuarios
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usuários', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usuários', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usuários', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usuários', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usuários', 6)

-- Clientes
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20402, 'Cadastro de Clientes', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20402, 'Cadastro de Clientes', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20402, 'Cadastro de Clientes', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20402, 'Cadastro de Clientes', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20402, 'Cadastro de Clientes', 6)

-- Alteração de Senha 
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10103, 'Alteração de Senhas', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10103, 'Alteração de Senhas', 5)

-- Marcas 
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20101, 'Cadastro de Marcas', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20101, 'Cadastro de Marcas', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20101, 'Cadastro de Marcas', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20101, 'Cadastro de Marcas', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20101, 'Cadastro de Marcas', 6)

-- Modelos
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20102, 'Cadastro de Modelos', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20102, 'Cadastro de Modelos', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20102, 'Cadastro de Modelos', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20102, 'Cadastro de Modelos', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20102, 'Cadastro de Modelos', 6)

-- Unidades
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20103, 'Cadastro de Unidades', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20103, 'Cadastro de Unidades', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20103, 'Cadastro de Unidades', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20103, 'Cadastro de Unidades', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20103, 'Cadastro de Unidades', 6)

-- Grupos
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20104, 'Cadastro de Grupos', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20104, 'Cadastro de Grupos', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20104, 'Cadastro de Grupos', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20104, 'Cadastro de Grupos', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20104, 'Cadastro de Grupos', 6)

-- Sub Grupos 
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20105, 'Cadastro de Sub-Grupos', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20105, 'Cadastro de Sub-Grupos', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20105, 'Cadastro de Sub-Grupos', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20105, 'Cadastro de Sub-Grupos', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20105, 'Cadastro de Sub-Grupos', 6)

-- Produto
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20200, 'Cadastro de Produtos', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20200, 'Cadastro de Produtos', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20200, 'Cadastro de Produtos', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20200, 'Cadastro de Produtos', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20200, 'Cadastro de Produtos', 6)
Go