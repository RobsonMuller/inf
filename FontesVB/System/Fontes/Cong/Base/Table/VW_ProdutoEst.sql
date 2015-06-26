Create View VW_ProdutoEst AS
	Select Produtos.Empresa, Produtos.Codigo, Produtos.Descricao, 
		Sum(EntradaEst.Qtd) AS EstEntrada,
		Sum(SaidaEst.Qtd) AS EstSaida,
		Sum(EntradaEst.Qtd - SaidaEst.Qtd) AS EstAtual
	From Produtos
	Left Join EntradaEst On (EntradaEst.Empresa = Produtos.Empresa And EntradaEst.CodProduto = Produtos.Codigo)
	Left Join SaidaEst On (SaidaEst.Empresa = Produtos.Empresa And SaidaEst.CodProduto = Produtos.Codigo)
	Where Produtos.IdSituacao = 1 -- Ativado