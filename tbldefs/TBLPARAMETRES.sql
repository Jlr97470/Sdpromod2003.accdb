CREATE TABLE [TBLPARAMETRES] (
  [ParType] VARCHAR (6),
  [ParCode] VARCHAR (255),
  [ParValeur] VARCHAR (100),
  [ParValide] BIT,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([ParType], [ParCode])
)
