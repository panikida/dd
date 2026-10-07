Public Sub CreateQueries()
    Dim db As DAO.Database
    Dim q As Variant
    Set db = CurrentDb

    On Error Resume Next
    For Each q In Array("Phones1", "ClientAddresses2", "Birthdays3", _
        "EmployeeBirthdays4", "BirthdaysByMonth5", "PhoneList6", _
        "CompletedOrders7", "AmountOver8", "AmountInRange9", "Managers10")
        db.QueryDefs.Delete q
    Next q
    On Error GoTo 0

    On Error Resume Next
    db.Execute "ALTER TABLE Employees ADD COLUMN Birthday DATETIME"
    On Error GoTo 0

    db.Execute "UPDATE Employees SET Birthday=#4/15/1985# WHERE EmployeeID=1"
    db.Execute "UPDATE Employees SET Birthday=#5/20/1990# WHERE EmployeeID=2"
    db.Execute "UPDATE Employees SET Birthday=#4/3/1988# WHERE EmployeeID=3"
    db.Execute "UPDATE Employees SET Birthday=#6/12/1975# WHERE EmployeeID=4"
    db.Execute "UPDATE Employees SET Birthday=#4/28/1992# WHERE EmployeeID=5"
    db.Execute "UPDATE Employees SET Birthday=#3/8/1980# WHERE EmployeeID=6"
    db.Execute "UPDATE Employees SET Birthday=#5/15/1985# WHERE EmployeeID=7"
    db.Execute "UPDATE Employees SET Birthday=#7/22/1978# WHERE EmployeeID=8"
    db.Execute "UPDATE Employees SET Birthday=#4/10/1990# WHERE EmployeeID=9"
    db.Execute "UPDATE Employees SET Birthday=#9/5/1982# WHERE EmployeeID=10"

    db.Execute "UPDATE Orders SET Amount=75000 WHERE OrderID=1"
    db.Execute "UPDATE Orders SET Amount=35000 WHERE OrderID=2"
    db.Execute "UPDATE Orders SET Amount=45000 WHERE OrderID=3"
    db.Execute "UPDATE Orders SET Amount=85000 WHERE OrderID=4"
    db.Execute "UPDATE Orders SET Amount=25000 WHERE OrderID=5"
    db.Execute "UPDATE Orders SET Amount=60000 WHERE OrderID=6"
    db.Execute "UPDATE Orders SET Amount=30000 WHERE OrderID=7"
    db.Execute "UPDATE Orders SET Amount=90000 WHERE OrderID=8"
    db.Execute "UPDATE Orders SET Amount=55000 WHERE OrderID=12"

    db.CreateQueryDef "Phones1", "SELECT LastName, FirstName, Phone FROM Employees"
    db.CreateQueryDef "ClientAddresses2", "SELECT CompanyName, Address, Phone FROM Clients ORDER BY CompanyName"
    db.CreateQueryDef "Birthdays3", "SELECT LastName, FirstName, Birthday FROM Employees"
    db.CreateQueryDef "EmployeeBirthdays4", "SELECT LastName, FirstName, Birthday FROM Employees"
    db.CreateQueryDef "BirthdaysByMonth5", "SELECT LastName, FirstName, Birthday FROM Employees WHERE Birthday LIKE [Enter date]"
    db.CreateQueryDef "PhoneList6", "SELECT LastName, FirstName, Phone FROM Employees WHERE LastName=[Enter last name]"

    db.CreateQueryDef "CompletedOrders7", _
        "SELECT Employees.LastName, Employees.FirstName, Clients.CompanyName, " & _
        "Orders.Completed, Orders.Amount, [Amount]*0.13 AS Tax, " & _
        "[Amount]-[Amount]*0.13 AS Profit " & _
        "FROM (Employees INNER JOIN Orders ON Employees.EmployeeID=Orders.EmployeeID) " & _
        "INNER JOIN Clients ON Clients.ClientID=Orders.ClientID " & _
        "WHERE Orders.Completed=True"

    db.CreateQueryDef "AmountOver8", _
        "SELECT Clients.CompanyName, Orders.Amount, Employees.LastName, Employees.FirstName, " & _
        "Orders.OrderDate, Orders.DueDate " & _
        "FROM (Employees INNER JOIN Orders ON Employees.EmployeeID=Orders.EmployeeID) " & _
        "INNER JOIN Clients ON Clients.ClientID=Orders.ClientID WHERE Orders.Amount>50000"

    db.CreateQueryDef "AmountInRange9", _
        "SELECT Clients.CompanyName, Orders.Amount, Employees.LastName, Employees.FirstName, " & _
        "Orders.OrderDate, Orders.DueDate " & _
        "FROM (Employees INNER JOIN Orders ON Employees.EmployeeID=Orders.EmployeeID) " & _
        "INNER JOIN Clients ON Clients.ClientID=Orders.ClientID " & _
        "WHERE Orders.Amount BETWEEN 20000 AND 50000"

    db.CreateQueryDef "Managers10", "SELECT LastName, FirstName, Position FROM Employees WHERE Position='Manager'"

    MsgBox "Queries created!", vbInformation
End Sub