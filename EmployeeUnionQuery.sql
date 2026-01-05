SELECT FolderID, EmployeeID, EmployeeFirstName, EmployeeLastName, IdentificationID, 'دائم' AS EmployeeType FROM PerEmployees
UNION ALL SELECT FolderID, EmployeeID, EmployeeFirstName, EmployeeLastName, IdentificationID, 'موقت' AS EmployeeType FROM TemEmployees;
