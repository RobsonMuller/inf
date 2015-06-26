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

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usuários', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usuários', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usuários', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usuários', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10101, 'Cadastro de Usuários', 6)

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20402, 'Cadastro de Clientes', 1)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20402, 'Cadastro de Clientes', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20402, 'Cadastro de Clientes', 3)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20402, 'Cadastro de Clientes', 4)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (20402, 'Cadastro de Clientes', 6)

INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10103, 'Alteração de Senhas', 2)
INSERT INTO dbo.GlbInterface (Codigo, Descricao, IdButton) Values (10103, 'Alteração de Senhas', 5)
Go