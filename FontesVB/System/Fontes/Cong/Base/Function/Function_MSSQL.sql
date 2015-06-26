Drop Function AccessSystem
Go

Create Function AccessSystem(@Emp Char(2), @NmUser Varchar(20), @PassWord Varchar(20)) Returns Varchar(255)
AS
Begin
	Declare @retTexto Varchar(255)
	
	Declare @retEmpAtiv Char(1)
	Declare @retUserAtiv Char(1)
	Declare @retNmUser Varchar(20)
	Declare @retPassWord Varchar(20)
	
	Declare CurEmp cursor for
	Select Empresas.NomeFantasia 
	From Empresas 
	Where Empresas.Codigo = @Emp
	
	Open CurEmp
	Fetch Next From CurEmp Into @retEmpAtiv
	if @retEmpAtiv  Is Null
		Begin
			Set @retTexto = 'Empresa n�o Licenciada! Verifique.'
			Return @retTexto
		End

	if @retEmpAtiv = 'N'
		Begin
			Set @retTexto = 'Empresa desativada! Verifique.'
			Return @retTexto
		End
	Close CurEmp

	Declare CurUser Cursor For
	Select Usuarios.Ativo, Usuarios.Nome, Usuarios.Senha
	From Usuarios 
	Where Usuarios.Empresa = @Emp 
	And Usuarios.Nome = @NmUser 
	
	Open CurUser
	Fetch Next From CurUser Into @retUserAtiv, @retNmUser, @retPassWord
	if @retUserAtiv IS Null
		Begin 
			Set @retTexto = 'Usu�rio n�o cadastrado! Verifique.'
			Return @retTexto
		End

	if @retUserAtiv = 'N'
		Begin
			Set @retTexto = 'Usu�rio desativado! Verifique.'
			Return @retTexto
		end
	
	if @retPassWord <> @PassWord
		Begin
			Set @retTexto = 'Senha inv�lida! Verifique.'
			Return @retTexto
		End
	else	
		Set @retTexto = 'OK'
	Close CurUSer

	return @rettexto
end


Select dbo.accesssystem('01','Master', 'Master')	



	--Retorno com Loop
	--While @@FETCH_STATUS = 0
	--BEGIN
	--END
