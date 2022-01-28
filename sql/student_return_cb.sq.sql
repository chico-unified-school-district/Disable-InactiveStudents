SELECT distinct STU.SC AS Sch,
STU.ID AS PermID,
STU.LN AS LastName,
STU.FN AS FirstName,
STU.PG AS Parentname,
STU.PEM AS ParentEMail,
STU.FW AS Fatherworkphone,
STU.MW AS Motherworkphone,
STU.GR AS Grade,
STU.SEM AS Mail,
 [PWA].[EM] AS ParentPortalEmail,
 DRI.BC AS Barcode,
DRI.SR AS SerialNumber,
 [DRA].[CD] AS [Code1],
[DRA].[CC] AS [Condition],
[DRA].[CO] AS [Comment],
CONVERT(varchar,DRA.DT,23) AS [IssuedDate],
STU.AD+', '+ STU.CY+', '+  STU.ST+' '+STU.ZC  AS [Address]
FROM STU INNER JOIN DRA ON STU.ID = DRA.ID AND DRA.ST = 'S' AND DRA.DEL = 0
INNER JOIN DRT ON DRA.RID = DRT.RID AND DRT.DEL = 0
INNER JOIN DRI ON DRA.RID = DRI.RID AND DRA.RIN = DRI.RIN
LEFT JOIN PWS  ON [STU].[ID] = [PWS].[ID]
LEFT JOIN PWA  ON [PWA].[AID] = [PWS].[AID]
WHERE  DRA.RD IS NULL
and [DRA].[CD] != 'S'
and [PWA].[TY] = 'P'
AND DRI.BC IS NOT NULL
AND STU.ID =  {0}