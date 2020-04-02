
IF NOT EXISTS
(
    SELECT *
    FROM sysobjects
    WHERE id = OBJECT_ID(N'[dbo].[tblShortcuts]')
)
    BEGIN
        CREATE TABLE tblShortcuts
        (ShortcutId   INT NOT NULL IDENTITY(1, 1) PRIMARY KEY, 
         ShortName    VARCHAR(MAX), 
         LongName     VARCHAR(MAX), 
         ShortcutCode INT
        );
END;