CREATE TABLE [dbo].[NecessaryFile]
(
	[Id] INT NOT NULL PRIMARY KEY IDENTITY(1,1),
	[FileName] NVARCHAR(256) NOT NULL,
	[Data] varbinary(max) NULL, 
    [UpdateTime] DATETIME NULL
)
