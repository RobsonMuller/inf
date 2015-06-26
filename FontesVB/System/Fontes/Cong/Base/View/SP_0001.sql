Drop View Vw_EstAtual
Go

Create View Vw_EstAtual As
	Select Produtos.Codigo AS IdProduto,
		Produtos.Descricao AS DescProduto,
		Produtos.CodBarras AS CodBarras,
		EstEntradaProd.IdEstEntrada,
		EstEntradaProd.VlrCusto,
		EstEntradaProd.Quantidade As EstEnt,
		ISNULL(EstSaidaProd.VlrVenda, 0),
		ISNULL(EstSaidaProd.Quantidade, 0) As EstSaid,
		(EstEntradaProd.Quantidade - ISNULL(EstSaidaProd.Quantidade, 0)) AS EstAtual
	From Produtos
	Inner Join EstEntradaProd On (EstEntradaProd.Empresa = Produtos.Empresa And EstEntradaProd.IdProduto = Produtos.Codigo)
	Left Join EstSaidaProd On (EstSaidaProd.Empresa = EstEntradaProd.Empresa And EstSaidaProd.IdProduto = EstEntradaProd.IdProduto)
Go