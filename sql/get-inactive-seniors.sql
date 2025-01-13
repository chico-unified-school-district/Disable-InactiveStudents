SELECT
    [STU].[ID] AS [permId]
    ,[STU].[ID] AS [employeeId]
    , [STU].[NID] AS [gmail]
    , [STU].[HSG] AS [completionStatus]
    , CONVERT(VARCHAR(10),[STU].[DG],101) AS [completionDate]
    , [STU].[TG] AS [statusTag]
FROM
    (SELECT [STU].*
    FROM STU
    WHERE [STU].DEL = 0) [STU]
WHERE
   [STU].[GR] = 12 AND [STU].[HSG] > ' '
    AND [STU].[DG] > DATEADD(month,-8,GETDATE())
;