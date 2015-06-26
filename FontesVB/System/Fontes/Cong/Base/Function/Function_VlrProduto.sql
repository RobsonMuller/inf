IF OBJECT_ID (N'dbo.VerificaLucroFornecedor', N'FN') IS NOT NULL
    DROP FUNCTION dbo.VerificaLucroFornecedor;
GO

Create Function dbo.VerificaLucroFornecedor (@VlrVenda Decimal(12,2), @VlrComp Decimal(12,2)) Returns Table AS
Return
(
   Select 
      CAST(ROUND((((@VlrVenda/@VlrComp)*100)-100),2) AS DECIMAL(5,2)) AS LucroPerc, 
      CAST(ROUND(((@VlrComp * (((@VlrVenda/@VlrComp)*100)-100))/100),2,1) AS DECIMAL(5,2)) AS LucroVlr
)
GO

SELECT (LucroVlr + 27.50) as valorvenda FROM VerificaLucroFornecedor(48.13, 27.50)
GO