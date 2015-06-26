if OBJECT_ID('dbo.GlbPermissoes') IS NOT NULL	
	Drop Table dbo.GlbPermissoes
Go
Create Table dbo.GlbPermissoes (
	Empresa Char(2) Not Null,
	IdUsuario Numeric(10) Not Null,
	IdInterface Numeric(10) Not Null,
	IdButton Numeric(10) Not Null
)
Alter Table dbo.GlbPermissoes Add Constraint PK_GLBPERMISSOES Primary Key (Empresa, IdUsuario, IdInterface, IdButton)
Alter Table dbo.GlbPermissoes Add Constraint FK_GLBPERMISSOES_EMPRESAS Foreign Key (Empresa) References Empresas (Codigo)
Alter Table dbo.GlbPermissoes Add Constraint FK_GLBPERMISSOES_GLBBUTTON Foreign Key (IdButton) References GlbButton (Codigo)
Go


