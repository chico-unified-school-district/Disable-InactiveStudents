SELECT DISTINCT STU.ID AS employeeID
 FROM STU 
  LEFT JOIN ENR ON STU.ID = ENR.ID 
 WHERE STU.TG != ''