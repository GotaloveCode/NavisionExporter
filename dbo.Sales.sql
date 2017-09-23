CREATE TABLE [dbo].[Sales] (
    [Id]          INT           IDENTITY (1, 1) NOT NULL,
    [Product]     VARCHAR (200) NULL,
    [Category]    VARCHAR (100) NULL,
    [Quantity]    FLOAT (53)    NULL,
    [Customer]    VARCHAR (100) NULL,
    [SalesPerson] NCHAR (100)   NULL,
    [Uploaded]    TINYINT       DEFAULT ((0)) NULL,
    PRIMARY KEY CLUSTERED ([Id] ASC)
);

