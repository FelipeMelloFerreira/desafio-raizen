CREATE TABLE DW_BI01.dbo.TB_DIESEL_SALES (
	YEAR_MONTH date NOT NULL,
	UF nvarchar(20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	PRODUTO nvarchar(40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	UNIT nvarchar(6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	VOLUME float NULL,
	[CREATE AT] datetime2(3) NOT NULL
);
CREATE NONCLUSTERED INDEX TB_DIESEL_SALES_PRODUTO_IDX ON DW_BI01.dbo.TB_DIESEL_SALES (PRODUTO, YEAR_MONTH);