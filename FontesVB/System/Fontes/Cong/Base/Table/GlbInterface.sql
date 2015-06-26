if OBJECT_ID('dbo.GlbInterface') IS NOT NULL	
	Drop From dbo.GlbInterface
Go

Create Table dbo.GlbInterface (
	Codigo Numeric(10) Not Null,
	Descricao Varchar(40) Not Null,
	IdButton Numeric(10) Not Null
)
Alter Table dbo.GlbInterface Add Constraint PK_GLBINTERFACE Primary Key (Codigo, IdButton)
Alter Table dbo.GlbInterface Add Constraint FK_GLBINTERFACE_GLBBUTTON Foreign Key (IdButton) References GlbButton (Codigo)
Go

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20101, 'Cadastro de Marcas', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20101, 'Cadastro de Marcas', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20101, 'Cadastro de Marcas', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20101, 'Cadastro de Marcas', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20101, 'Cadastro de Marcas', 6)

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usuários', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usuários', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usuários', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usuários', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usuários', 6)

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20200, 'Cadastro de Produtos', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20200, 'Cadastro de Produtos', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20200, 'Cadastro de Produtos', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20200, 'Cadastro de Produtos', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20200, 'Cadastro de Produtos', 6)

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20401, 'Cadastro de Fornecedores', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20401, 'Cadastro de Fornecedores', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20401, 'Cadastro de Fornecedores', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20401, 'Cadastro de Fornecedores', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20401, 'Cadastro de Fornecedores', 6)

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20402, 'Cadastro de Clientes', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20402, 'Cadastro de Clientes', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20402, 'Cadastro de Clientes', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20402, 'Cadastro de Clientes', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20402, 'Cadastro de Clientes', 6)

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20501, 'Envio de E-mail', 8)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20502, 'Recebimento de E-mail', 8)

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (40100, 'Configuração de E-mail', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (40100, 'Configuração de E-mail', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (40100, 'Configuração de E-mail', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (40100, 'Configuração de E-mail', 6)

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (30101, 'Relatório de Marcas', 99)

Go