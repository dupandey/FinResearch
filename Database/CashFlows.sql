USE [FinResearch]
GO

/****** Object:  Table [dbo].[CashFlows]    Script Date: 3/22/2020 1:34:11 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[CashFlows](
	[CashFlowId] [bigint] IDENTITY(1,1) NOT NULL,
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
 CONSTRAINT [PK_CashFlows] PRIMARY KEY CLUSTERED 
(
	[CashFlowId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

ALTER TABLE [dbo].[CashFlows]  WITH CHECK ADD  CONSTRAINT [FK_CashFlows_Category] FOREIGN KEY([CategoryId])
REFERENCES [dbo].[Category] ([CategoryId])
GO

ALTER TABLE [dbo].[CashFlows] CHECK CONSTRAINT [FK_CashFlows_Category]
GO

ALTER TABLE [dbo].[CashFlows]  WITH CHECK ADD  CONSTRAINT [FK_CashFlows_Company] FOREIGN KEY([CompanyId])
REFERENCES [dbo].[Company] ([CompanyId])
GO

ALTER TABLE [dbo].[CashFlows] CHECK CONSTRAINT [FK_CashFlows_Company]
GO

