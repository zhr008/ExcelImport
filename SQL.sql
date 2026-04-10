/*


USE [TEST]
GO

/****** Object:  Table [dbo].[ExcelImportRow]    Script Date: 2026-04-10 10:37:02 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ExcelImportRow](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[TestId] [varchar](50) NOT NULL,
	[MeasuringSide] [varchar](50) NULL,
	[TaskNumber] [varchar](50) NULL,
	[Description] [varchar](50) NULL,
	[TestTime] [datetime] NULL,
	[Operator] [varchar](50) NULL,
	[IsPass] [int] NULL,
	[Wavelength1310_IL] [decimal](12, 6) NULL,
	[Wavelength1310_RL] [decimal](12, 6) NULL,
	[Wavelength1550_IL] [decimal](12, 6) NULL,
	[Wavelength1550_RL] [decimal](12, 6) NULL
) ON [PRIMARY]
GO

USE [TEST]
GO

/****** Object:  Table [dbo].[ExcelImportCell]    Script Date: 2026-04-10 10:36:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ExcelImportCell](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[TestId] [varchar](50) NOT NULL,
	[MeasuringSide] [varchar](50) NULL,
	[TaskNumber] [varchar](50) NULL,
	[Description] [varchar](50) NULL,
	[TestTime] [datetime] NULL,
	[Operator] [varchar](50) NULL,
	[IsPass] [int] NULL,
	[Wavelength1310_IL] [decimal](12, 6) NULL,
	[Wavelength1310_RL] [decimal](12, 6) NULL
) ON [PRIMARY]
GO

*/


SELECT * from 
--truncate table
ExcelImportCell

SELECT * from 
--truncate table
ExcelImportRow



















