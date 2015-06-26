-- Teste de Desenvolimento

DELETE Usuarios 
DELETE FROM EmpresasContatos
DELETE FROM Empresas
GO

INSERT INTO dbo.Empresas (Codigo, RazaoSocial, NomeFantasia, CPFCNPJ, FoneCom, RamalCom, FoneCel, FoneRes, DtFundacao, DtMovimento, Email, WebSite, Obs, Situacao)
VALUES ('01', 'Infinity Sistemas de Informação', 'Infinity', '02182135065', '05199804745', '0', '05199804745', '05135823041', GetDate(), {d '2015-06-21'}, 'infinity.gerencia@gmail.com', NULL, 'Empresa para teste', 'A')
GO

INSERT INTO dbo.EmpresasContatos (Empresa, Codigo, Nome, Setor, Fone, Ramal, Email, Obs) 
VALUES ('01', 1, 'Robson Muller', 'Diretoria', '05199804745', '0', 'infinity.gerencia@gmail.com', 'Primeiro Contato')
GO

INSERT INTO dbo.Usuarios (Empresa, Codigo, Nome, Usuario, Senha, NivelAcesso, Situacao)
VALUES ('01', 1, 'Desenvolvimento Infinity', 'master', 'EB0A191797624DD3A48FA681D3061212', 'A', 'A') -- Senha master (Criptografada)
GO
