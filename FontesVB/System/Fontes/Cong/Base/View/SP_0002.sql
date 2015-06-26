Drop View Vw_Prod_Ativo
Go

Create View Vw_Prod_Ativo AS 
	Select Produtos.Codigo AS IdProd, 
		Produtos.Descricao AS DescProd,
		Produtos.DtValid,
		Produtos.EstLimite,
		Produtos.ControlaEst,
		Produtos.VlrCompra,
		Produtos.CodBarras,
		Marcas.Codigo AS IdMarca,
		Marcas.Descricao AS DescMarca,
		Modelos.Codigo AS IdModelo,
		Modelos.Descricao AS DescModelo,
		Unidades.Codigo AS IdUnidade,
		Unidades.Descricao AS DescUnidade,
		Categorias.Codigo AS IdCategoria,
		Categorias.Descricao As DescCategoria
		Fornecedores.Codigo As IdForn,
		Fornecedores.RazaoSocial AS RazaoSocialForn		
	From Produtos
	Inner Join Empresas (Empresas.Empresa = Produtos.Empresa)
	Inner Join ProdutosForn (ProdutosForn.Empresa = Produtos.Empresa And ProdutosForn.CodProduto = Produtos.Codigo)
	Inner Join Fornecedores (Fornecedores.Empresa = ProdutosForn.Empresa And Fornecedores.Codigo And ProdutosForn.CodFornecedor)
	Left Join Marcas On (Marcas.Empresa = Produtos.Empresa And Marcas.Codigo = Produtos.IdMarca)
	Left Join Modelos On (Modelos.Empresa = Produtos.Empresa And Modelos.Codigo = Produtos.IdMarca)
	Left Join Unidades On (Unidades.Empresa = Produtos.Empresa And Unidades.Codigo = Produtos.IdMarca)
	Left Join Categoria On (Categorias.Empresa = Produtos.Empresa And Categorias.Codigo = Produtos.IdMarca)
	Where Produtos.Ativo = 'S'