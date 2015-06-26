If Object_ID('dbo.Produtos') Is Not Null 
	Drop Table dbo.Produtos
Go 
Create Table dbo.Produtos (
	Empresa Char(2) Not Null,
	Codigo Numeric(10) Not Null,
	CodBarra Numeric(13) Not Null,
	Descricao Varchar(80) Not Null,
	Abreviatura Varchar(40) Null,
	CodGrupo Numeric(10) Null,
	CodSubGrupo Numeric(10) Null,
	CodMarca Numeric(10) Null,
	CodModelo Numeric(10) Null,
	CodUnidade Numeric(10) Not Null,
	Obs Varchar(255) Null,
	IdSituacao Char(1) Not Null,
	LucroMin Numeric(5, 2) Not Null,
	IdLucroMin Char(1) Not Null,
	DtCad DateTime Not Null,
	CodUsuario Numeric(10) Not Null,
	DtUltAlt DateTime Null,
	CodUsuarioUltAlt Numeric(10) Null,
	IdControlaEst Char(1) Not Null,
	EstoqueMin Numeric(5) Null,
	IdVenderSemEst Char(1) Not Null,
	IdTributacao Char(1) Not Null,
	ICMS Numeric(13, 2) Null,
	IdICMS Char(1) Not Null,
	PISCOFINS Numeric(13, 2) Null,
	IdPISCOFINS Char(1) Not Null,
	IPI Numeric(13, 2) Null,
	IdIPI Char(1) Not Null,
	Tributos Numeric(13, 2) Null,
	Frete Numeric(13, 2) Null, 
	IdFrete Char(1) Not Null,
	Comissao Numeric(13, 2) Null,
	IdComissao Char(1) Not Null,
	Margem Numeric(13, 2) Null,
	IdMargem Char(1) Not Null,
	Custos Numeric(13, 2) Null,
	IdCustos Char(1) Not Null,
	ValorVenda Numeric(13, 2) Null
)
Alter Table dbo.Produtos Add Constraint PK_PRODUTOS Primary Key (Empresa, Codigo)

Alter Table dbo.Produtos Add Constraint FK_PRODUTOS_EMPRESAS Foreign Key (Empresa) References dbo.Empresas (Codigo)
Alter Table dbo.Produtos Add Constraint FK_PRODUTOS_GRUPO Foreign Key (CodGrupo) References dbo.Grupos (Codigo)
Alter Table dbo.Produtos Add Constraint FK_PRODUTOS_SUBGRUPO Foreign Key (CodSubGrupo) References dbo.SubGrupos (Codigo)
Alter Table dbo.Produtos Add Constraint FK_PRODUTOS_MARCA Foreign Key (CodMarca) References dbo.Marcas (Codigo)
Alter Table dbo.Produtos Add Constraint FK_PRODUTOS_MODELO Foreign Key (CodModelo) References dbo.Modelos (Codigo)
Alter Table dbo.Produtos Add Constraint FK_PRODUTOS_UNIDADE Foreign Key (CodUnidade) References dbo.Unidades (Codigo)

Alter Table dbo.Produtos Add Constraint CK_PRODUTOS_IDSITUACAO Check (IdSituacao IN (0, 1)) --0 = Desativado, 1 = Ativado
Alter Table dbo.Produtos Add Constraint CK_PRODUTOS_IDLUCROMIN Check (IdLucroMin IN (0, 1)) --0 = %, 1 = R$
Alter Table dbo.Produtos Add Constraint CK_PRODUTOS_IDCONTROLAEST Check (IdControlaEst IN ('S', 'N')) 
Alter Table dbo.Produtos Add Constraint CK_PRODUTOS_IDVENDERSEMEST Check (IdVenderSemEst IN ('S', 'N')) 
Alter Table dbo.Produtos Add Constraint CK_PRODUTOS_IDTRIBUTACAO Check (IdTributacao IN (0, 1))  -- 0 = Isento, 1 = Tributado
Alter Table dbo.Produtos Add Constraint CK_PRODUTOS_IDICMS Check (IdICMS IN (0, 1))  --0 = %, 1 = R$
Alter Table dbo.Produtos Add Constraint CK_PRODUTOS_IDPISCOFINS Check (IdPISCOFINS IN (0, 1))  --0 = %, 1 = R$
Alter Table dbo.Produtos Add Constraint CK_PRODUTOS_IDIPI Check (IdIPI IN (0, 1))  --0 = %, 1 = R$
Alter Table dbo.Produtos Add Constraint CK_PRODUTOS_IDCOMISSAO Check (IdComissao IN (0, 1))  --0 = %, 1 = R$
Alter Table dbo.Produtos Add Constraint CK_PRODUTOS_IDMARGEM Check (IdMargem IN (0, 1))  --0 = %, 1 = R$
Alter Table dbo.Produtos Add Constraint CK_PRODUTOS_IDCUSTOS Check (IdCustos IN (0, 1))  --0 = %, 1 = R$
Go