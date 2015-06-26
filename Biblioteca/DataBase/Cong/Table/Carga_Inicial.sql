-- Carga Inicial
INSERT INTO Empresas (Codigo, RazaoSocial, CPFCNPJ, DataFundacao, Situacao) VALUES ('01', 'Infinity Sistemas de Informação', '02182135065', GetDate(), 'S')

INSERT INTO Usuarios (Empresa, Codigo, Nome, Usuario, Senha, NivelAcesso, Situacao) VALUES ('01', 1, 'Master', 'Master', 'master', 1, 'S')

