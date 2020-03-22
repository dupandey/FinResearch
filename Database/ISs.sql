CREATE TABLE [dbo].[ISs](
	[ISId] [bigint] IDENTITY(1,1) NOT NULL,
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
 CONSTRAINT [PK_ISs] PRIMARY KEY CLUSTERED 
(
	[ISId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

ALTER TABLE [dbo].[ISs]  WITH CHECK ADD  CONSTRAINT [FK_ISs_Category] FOREIGN KEY([CategoryId])
REFERENCES [dbo].[Category] ([CategoryId])
GO

ALTER TABLE [dbo].[ISs] CHECK CONSTRAINT [FK_ISs_Category]
GO

ALTER TABLE [dbo].[ISs]  WITH CHECK ADD  CONSTRAINT [FK_ISs_Company] FOREIGN KEY([CompanyId])
REFERENCES [dbo].[Company] ([CompanyId])
GO

ALTER TABLE [dbo].[ISs] CHECK CONSTRAINT [FK_ISs_Company]
GO


