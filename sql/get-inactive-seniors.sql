WITH RankedStudents AS (
    SELECT
        [STU].[ID]
        , [STU].[NID] AS [gmail]
        , [STU].[HSG] AS [completionStatus]
        , CONVERT(VARCHAR(10), [STU].[DG], 101) AS [completionDate]
        , [STU].[TG] AS [statusTag]
        , ROW_NUMBER() OVER (
            PARTITION BY [STU].[ID]
            ORDER BY [STU].[DG] DESC, [STU].[NID] DESC
          ) AS [rn]
    FROM [STU]
    WHERE [STU].[DEL] = 0
      AND [STU].[GR] = 12
      AND [STU].[HSG] > ' '
      AND [STU].[DG] IS NOT NULL
)
SELECT
    [ID]
    , [gmail]
    , [completionStatus]
    , [completionDate]
    , [statusTag]
FROM RankedStudents
WHERE [rn] = 1;