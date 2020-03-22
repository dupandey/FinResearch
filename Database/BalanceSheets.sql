USE [FinResearch]
GO

/****** Object:  Table [dbo].[BalanceSheets]    Script Date: 3/22/2020 1:33:21 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[BalanceSheets](
	[BalanceSheetId] [bigint] IDENTITY(1,1) NOT NULL,
	[CompanyId] [int] NOT NULL,
	[CategoryId] [int] NOT NULL,
	[Payload] [nvarchar](max) NULL,
	[IsActive] [bit] NOT NULL,
	[CreatedDate] [datetime] NULL,
	[ModifiedDate] [datetime] NULL,
	[FileName] [nvarchar](300) NULL,
	[UserId] [nvarchar](128) NULL,
	[IsAdmin] [bit] NULL,
	[FileVersion] [int] NULL,
 CONSTRAINT [PK_BalanceSheets] PRIMARY KEY CLUSTERED 
(
	[BalanceSheetId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

ALTER TABLE [dbo].[BalanceSheets]  WITH CHECK ADD  CONSTRAINT [FK_BalanceSheets_Category] FOREIGN KEY([CategoryId])
REFERENCES [dbo].[Category] ([CategoryId])
GO

ALTER TABLE [dbo].[BalanceSheets] CHECK CONSTRAINT [FK_BalanceSheets_Category]
GO

ALTER TABLE [dbo].[BalanceSheets]  WITH CHECK ADD  CONSTRAINT [FK_BalanceSheets_Company] FOREIGN KEY([CompanyId])
REFERENCES [dbo].[Company] ([CompanyId])
GO

ALTER TABLE [dbo].[BalanceSheets] CHECK CONSTRAINT [FK_BalanceSheets_Company]
GO


