IF object_id('dbo.GLBButton') IS NOT NULL
	DROP TABLE dbo.GLBButton 
GO

CREATE TABLE dbo.GLBButton (
	Codigo Numeric(10) Not Null,
	Descricao Varchar(20) Not Null
)
ALTER TABLE dbo.GLBButton ADD CONSTRAINT PK_GLBButton PRIMARY KEY (Codigo)
GO

INSERT INTO dbo.GlbButton (Codigo, Descricao) Values (1, 'Novo')
INSERT INTO dbo.GlbButton (Codigo, Descricao) Values (2, 'Consultar')
INSERT INTO dbo.GlbButton (Codigo, Descricao) Values (3, 'Inserir')
INSERT INTO dbo.GlbButton (Codigo, Descricao) Values (4, 'Alterar')
INSERT INTO dbo.GlbButton (Codigo, Descricao) Values (5, 'Salvar')
INSERT INTO dbo.GlbButton (Codigo, Descricao) Values (6, 'Excluir')  
INSERT INTO dbo.GlbButton (Codigo, Descricao) Values (7, 'Listar')
INSERT INTO dbo.GlbButton (Codigo, Descricao) Values (8, 'Enviar')
INSERT INTO dbo.GlbButton (Codigo, Descricao) Values (99, 'Relatórios')
GO