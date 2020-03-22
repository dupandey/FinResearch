CREATE TABLE [dbo].[ISs](
	[ISId] [bigint] IDENTITY(1,1) NOT NULL,
	[CompanyId] [int] NOT NULL,
	[CategoryId] [int] NOT NULL,
	[Payload] [nvarchar](max) NULL,
	[IsActive] [bit] NOT NULL,
	[CreatedDate] [datetime] NULL,
	[ModifiedDate] [datetime] NULL,
 CONSTRAINT [PK_ISs] PRIMARY KEY CLUSTERED 
(
	[ISId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO