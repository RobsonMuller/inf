If OBJECT_ID('dbo.ConfigEmail') IS NOT NULL	
	Drop Table dbo.ConfigEmail
GO

Create Table dbo.ConfigEmail (
	-- Parametros
	Empresa Char(2) Not Null,
	CodUser Numeric(10) Not Null,

	-- Envio
	SMTPServerPorta Numeric(5) Not Null,
	SMTPServer Varchar(80) Not Null,
	SMTPAuthenticateSSL Char(1) Not Null,
	SMTPAuthenticateServer Char(1) Not Null,
	
	-- Recebimento 
	POPServerPorta Numeric(5) Not Null, 
	POPServer Varchar(80) Not Null, 
	POPAuthenticateSSL Char(1) Not Null, 
	POPServerCopy Char(1) Not Null,
	POPTypeAuthenticate Char(1) Not Null,
	POPProtocolRead Char(1) Not Null, 
	
	-- Autenticacao
	UserName Varchar(80) Not Null,
	PassWord Varchar(40) Not Null,
	Email Varchar(255) Not Null,
	SendErro Char(1) Not Null
)

Alter Table dbo.ConfigEmail Add Constraint PK_CONFIGEMAIL Primary Key (Empresa, CodUser)
Go