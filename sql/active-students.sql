SELECT DISTINCT STU.ID AS employeeID
 FROM STU 
  LEFT JOIN ENR ON STU.ID = ENR.ID
 WHERE ( (STU.del = 0) OR (STU.del IS NULL) ) 
  AND ( STU.tg = ' ' ) 
  AND STU.SC != 999