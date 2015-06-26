If Object_ID('dbo.Usuarios') Is Not Null 
	Drop Table dbo.Usuarios
Go 
	
Create Table dbo.Usuarios (
	Empresa Char(2) Not Null,
	Codigo Numeric(10) Not Null,
	Nome Varchar(40) Not Null,
	Usuario Varchar(20) Not Null,
	Senha Varchar(40) Not Null,
	IdSituacao Char(1) Not Null,
	IdPerfil Char(1) Not Null,
	IdAlterarSenha Char(1) Not Null
)
Alter Table dbo.Usuarios Add Constraint PK_USUARIOS Primary Key (Empresa, Codigo)
Alter Table dbo.Usuarios Add Constraint CK_USUARIOS_IDSITUACAO Check (IdSituacao IN (1, 2, 3)) 
Alter Table dbo.Usuarios Add Constraint CK_USUARIOS_IDPERFIL Check (IdPerfil IN (1, 2))
Alter Table dbo.Usuarios Add Constraint CK_USUARIOS_IDALTERARSENHA Check (IdAlterarSenha IN ('S', 'N'))
GO

-- Carga
Insert Into dbo.Usuarios (Empresa, Codigo, Nome, Usuario, Senha, IdSituacao, IdPerfil, IdAlterarSenha) 
Values ('01', 1, 'Infinity Sistema de Informação', 'Master', 'f03bde11d261f185cbacfa32c1c6538c', 1, 1, 'N') -- Senha: Master
GO

-- Situacao
-- Ativo = 1
-- Inativo = 2
-- Bloqueado = 3

-- Perfil
-- Administrador = 1
-- Usuario = 2

If Object_ID('dbo.UsuariosExcluidos') Is Not Null 
	Drop Table dbo.UsuariosExcluidos
Go 

Create Table dbo.UsuariosExcluidos (
	Empresa Char(2) Not Null,
	Codigo Numeric(10) Not Null,
	Nome Varchar(40) Not Null,
	Usuario Varchar(20) Not Null,
	IdSituacao Char(1) Not Null,
	IdPerfil Char(1) Not Null,
	DtHrExclusao DateTime Not Null,
	CodUsuarioExclusao Numeric(10) Not Null
)
Alter Table dbo.UsuariosExcluidos Add Constraint PK_USUARIOSEXCLUIDOS Primary Key (Empresa, Codigo)
Go