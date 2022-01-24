SELECT distinct STU.SC School,
STU.ID PermID,
STU.LN LasstName,
STU.FN FirstName,
STU.PG Parentname,
STU.PEM ParentEMail,
STU.FW  Fatherworkphone,
STU.MW Motherworkphone,
 [PWA].[EM] ParentportalEmail,
 DRI.BC Barcode,
DRI.SR serial,
 [DRA].[CD] AS [Code1],
[DRA].[CC] AS [Condition],
[DRA].[CO] AS [Comment],
DRA.DT AS [IssuedDate],
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