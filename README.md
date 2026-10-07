Option Compare Database
Option Explicit

Public Sub CreateAll()
    Dim db As DAO.Database
    Dim rel As DAO.Relation
    Dim q As Variant
    Set db = CurrentDb

    ' ---------- CLEANUP ----------
    On Error Resume Next
    For Each q In Array("Phones1", "ClientAddresses2", "Birthdays3", _
        "EmployeeBirthdays4", "BirthdaysByMonth5", "PhoneList6", _
        "CompletedOrders7", "AmountOver8", "AmountInRange9", "Managers10")
        db.QueryDefs.Delete q
    Next q
    db.Relations.Delete "EmployeesOrders"
    db.Relations.Delete "ClientsOrders"
    db.Execute "DROP TABLE Orders"
    db.Execute "DROP TABLE Clients"
    db.Execute "DROP TABLE Employees"
    On Error GoTo 0

    ' ---------- TABLES ----------
    db.Execute "CREATE TABLE Employees (" & _
        "[EmployeeID] COUNTER PRIMARY KEY, " & _
        "[LastName] TEXT(50), [FirstName] TEXT(50), [MiddleName] TEXT(50), " & _
        "[Position] TEXT(50), [Phone] TEXT(20), [Address] TEXT(100), " & _
        "[Birthday] DATETIME)", dbFailOnError

    db.Execute "CREATE TABLE Clients (" & _
        "[ClientID] COUNTER PRIMARY KEY, " & _
        "[CompanyName] TEXT(100), [Address] TEXT(100), [Phone] TEXT(20), " & _
        "[Fax] TEXT(20), [Email] TEXT(100), [Notes] MEMO)", dbFailOnError

    db.Execute "CREATE TABLE Orders (" & _
        "[OrderID] COUNTER PRIMARY KEY, " & _
        "[ClientID] LONG, [EmployeeID] LONG, " & _
        "[OrderDate] DATETIME, [DueDate] DATETIME, " & _
        "[Amount] CURRENCY, [Completed] YESNO)", dbFailOnError

    ' ---------- RELATIONS ----------
    Set rel = db.CreateRelation("EmployeesOrders", "Employees", "Orders", dbRelationUpdateCascade Or dbRelationDeleteCascade)
    rel.Fields.Append rel.CreateField("EmployeeID")
    rel.Fields("EmployeeID").ForeignName = "EmployeeID"
    db.Relations.Append rel

    Set rel = db.CreateRelation("ClientsOrders", "Clients", "Orders", dbRelationUpdateCascade Or dbRelationDeleteCascade)
    rel.Fields.Append rel.CreateField("ClientID")
    rel.Fields("ClientID").ForeignName = "ClientID"
    db.Relations.Append rel

    ' ---------- LOOKUPS ----------
    SetFieldLookup db, "Orders", "EmployeeID", _
        "SELECT [EmployeeID], [LastName] & ' ' & [FirstName] AS FN FROM Employees ORDER BY [LastName];"
    SetFieldLookup db, "Orders", "ClientID", _
        "SELECT [ClientID], [CompanyName] FROM Clients ORDER BY [CompanyName];"

    ' ---------- EMPLOYEES ----------
    db.Execute "INSERT INTO Employees ([LastName],[FirstName],[MiddleName],[Position],[Phone],[Address],[Birthday]) VALUES ('Ivanov','Ivan','Ivanovich','Director','111-11-11','Lenina 1',#4/15/1985#)"
    db.Execute "INSERT INTO Employees ([LastName],[FirstName],[MiddleName],[Position],[Phone],[Address],[Birthday]) VALUES ('Petrov','Petr','Petrovich','Economist','222-22-22','Pushkina 2',#5/20/1990#)"
    db.Execute "INSERT INTO Employees ([LastName],[FirstName],[MiddleName],[Position],[Phone],[Address],[Birthday]) VALUES ('Sidorova','Anna','Sergeevna','Accountant','333-33-33','Gogolya 3',#4/3/1988#)"
    db.Execute "INSERT INTO Employees ([LastName],[FirstName],[MiddleName],[Position],[Phone],[Address],[Birthday]) VALUES ('Kuznetsov','Dmitry','Alekseevich','Manager','444-44-44','Chekhova 4',#6/12/1975#)"
    db.Execute "INSERT INTO Employees ([LastName],[FirstName],[MiddleName],[Position],[Phone],[Address],[Birthday]) VALUES ('Smirnova','Elena','Viktorovna','Manager','555-55-55','Tolstogo 5',#4/28/1992#)"
    db.Execute "INSERT INTO Employees ([LastName],[FirstName],[MiddleName],[Position],[Phone],[Address],[Birthday]) VALUES ('Popov','Alexey','Nikolaevich','Economist','666-66-66','Dostoevskogo 6',#3/8/1980#)"
    db.Execute "INSERT INTO Employees ([LastName],[FirstName],[MiddleName],[Position],[Phone],[Address],[Birthday]) VALUES ('Vasilyeva','Olga','Ivanovna','Accountant','777-77-77','Turgeneva 7',#5/15/1985#)"
    db.Execute "INSERT INTO Employees ([LastName],[FirstName],[MiddleName],[Position],[Phone],[Address],[Birthday]) VALUES ('Sokolov','Andrey','Vladimirovich','Manager','888-88-88','Nekrasova 8',#7/22/1978#)"
    db.Execute "INSERT INTO Employees ([LastName],[FirstName],[MiddleName],[Position],[Phone],[Address],[Birthday]) VALUES ('Mikhailova','Tatyana','Petrovna','Manager','999-99-99','Krylova 9',#4/10/1990#)"
    db.Execute "INSERT INTO Employees ([LastName],[FirstName],[MiddleName],[Position],[Phone],[Address],[Birthday]) VALUES ('Fedorov','Sergey','Dmitrievich','Manager','000-00-00','Lermontova 10',#9/5/1982#)"

    ' ---------- CLIENTS ----------
    db.Execute "INSERT INTO Clients ([CompanyName],[Address],[Phone],[Fax],[Email],[Notes]) VALUES ('Romashka LLC','Sadovaya 1','101-01-01','101-01-02','romashka@mail.ru','Regular')"
    db.Execute "INSERT INTO Clients ([CompanyName],[Address],[Phone],[Fax],[Email],[Notes]) VALUES ('Vektor JSC','Lugovaya 2','202-02-02','202-02-03','vector@mail.ru','')"
    db.Execute "INSERT INTO Clients ([CompanyName],[Address],[Phone],[Fax],[Email],[Notes]) VALUES ('Sidorov IE','Polevaya 3','303-03-03','','sidorov@mail.ru','')"
    db.Execute "INSERT INTO Clients ([CompanyName],[Address],[Phone],[Fax],[Email],[Notes]) VALUES ('TechnoService LLC','Zavodskaya 4','404-04-04','404-04-05','tech@mail.ru','')"
    db.Execute "INSERT INTO Clients ([CompanyName],[Address],[Phone],[Fax],[Email],[Notes]) VALUES ('Energia OJSC','Severnaya 5','505-05-05','505-05-06','energy@mail.ru','')"
    db.Execute "INSERT INTO Clients ([CompanyName],[Address],[Phone],[Fax],[Email],[Notes]) VALUES ('StroyMaster LLC','Yuzhnaya 6','606-06-06','606-06-07','stroy@mail.ru','')"
    db.Execute "INSERT INTO Clients ([CompanyName],[Address],[Phone],[Fax],[Email],[Notes]) VALUES ('Alfa JSC','Zapadnaya 7','707-07-07','707-07-08','alfa@mail.ru','')"
    db.Execute "INSERT INTO Clients ([CompanyName],[Address],[Phone],[Fax],[Email],[Notes]) VALUES ('Gamma LLC','Vostochnaya 8','808-08-08','808-08-09','gamma@mail.ru','')"
    db.Execute "INSERT INTO Clients ([CompanyName],[Address],[Phone],[Fax],[Email],[Notes]) VALUES ('Petrova IE','Tsentralnaya 9','909-09-09','','petrova@mail.ru','')"
    db.Execute "INSERT INTO Clients ([CompanyName],[Address],[Phone],[Fax],[Email],[Notes]) VALUES ('Delta LLC','Shkolnaya 10','010-10-10','010-10-11','delta@mail.ru','')"

    ' ---------- ORDERS ----------
    db.Execute "INSERT INTO Orders ([ClientID],[EmployeeID],[OrderDate],[DueDate],[Amount],[Completed]) VALUES (1,4,#1/10/2024#,#1/20/2024#,75000,True)"
    db.Execute "INSERT INTO Orders ([ClientID],[EmployeeID],[OrderDate],[DueDate],[Amount],[Completed]) VALUES (2,5,#1/12/2024#,#1/25/2024#,35000,False)"
    db.Execute "INSERT INTO Orders ([ClientID],[EmployeeID],[OrderDate],[DueDate],[Amount],[Completed]) VALUES (3,8,#1/15/2024#,#1/30/2024#,45000,True)"
    db.Execute "INSERT INTO Orders ([ClientID],[EmployeeID],[OrderDate],[DueDate],[Amount],[Completed]) VALUES (4,9,#1/18/2024#,#2/1/2024#,85000,False)"
    db.Execute "INSERT INTO Orders ([ClientID],[EmployeeID],[OrderDate],[DueDate],[Amount],[Completed]) VALUES (5,10,#1/20/2024#,#2/5/2024#,25000,True)"
    db.Execute "INSERT INTO Orders ([ClientID],[EmployeeID],[OrderDate],[DueDate],[Amount],[Completed]) VALUES (6,4,#1/22/2024#,#2/10/2024#,60000,False)"
    db.Execute "INSERT INTO Orders ([ClientID],[EmployeeID],[OrderDate],[DueDate],[Amount],[Completed]) VALUES (7,5,#1/25/2024#,#2/12/2024#,30000,True)"
    db.Execute "INSERT INTO Orders ([ClientID],[EmployeeID],[OrderDate],[DueDate],[Amount],[Completed]) VALUES (8,8,#2/1/2024#,#2/15/2024#,90000,False)"
    db.Execute "INSERT INTO Orders ([ClientID],[EmployeeID],[OrderDate],[DueDate],[Amount],[Completed]) VALUES (9,9,#2/3/2024#,#2/18/2024#,8000,True)"
    db.Execute "INSERT INTO Orders ([ClientID],[EmployeeID],[OrderDate],[DueDate],[Amount],[Completed]) VALUES (10,10,#2/5/2024#,#2/20/2024#,33000,False)"
    db.Execute "INSERT INTO Orders ([ClientID],[EmployeeID],[OrderDate],[DueDate],[Amount],[Completed]) VALUES (1,4,#2/10/2024#,#2/25/2024#,12000,True)"
    db.Execute "INSERT INTO Orders ([ClientID],[EmployeeID],[OrderDate],[DueDate],[Amount],[Completed]) VALUES (2,5,#2/12/2024#,#2/28/2024#,55000,False)"
    db.Execute "INSERT INTO Orders ([ClientID],[EmployeeID],[OrderDate],[DueDate],[Amount],[Completed]) VALUES (3,8,#2/15/2024#,#3/1/2024#,17000,True)"
    db.Execute "INSERT INTO Orders ([ClientID],[EmployeeID],[OrderDate],[DueDate],[Amount],[Completed]) VALUES (4,9,#2/18/2024#,#3/5/2024#,29000,False)"
    db.Execute "INSERT INTO Orders ([ClientID],[EmployeeID],[OrderDate],[DueDate],[Amount],[Completed]) VALUES (5,10,#2/20/2024#,#3/10/2024#,41000,True)"

    ' ---------- QUERIES ----------
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
        "INNER JOIN Clients ON Clients.ClientID=Orders.ClientID WHERE Orders.Completed=True"

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

    MsgBox "ALL DONE!", vbInformation
End Sub

Private Sub SetFieldLookup(db As DAO.Database, strTable As String, strField As String, strRowSource As String)
    Dim fld As DAO.Field
    Set fld = db.TableDefs(strTable).Fields(strField)
    On Error Resume Next
    fld.Properties.Delete "DisplayControl"
    fld.Properties.Delete "RowSourceType"
    fld.Properties.Delete "RowSource"
    fld.Properties.Delete "BoundColumn"
    fld.Properties.Delete "ColumnCount"
    fld.Properties.Delete "ColumnWidths"
    fld.Properties.Delete "LimitToList"
    On Error GoTo 0
    fld.Properties.Append fld.CreateProperty("DisplayControl", dbInteger, 111)
    fld.Properties.Append fld.CreateProperty("RowSourceType", dbText, "Table/Query")
    fld.Properties.Append fld.CreateProperty("RowSource", dbText, strRowSource)
    fld.Properties.Append fld.CreateProperty("BoundColumn", dbInteger, 1)
    fld.Properties.Append fld.CreateProperty("ColumnCount", dbInteger, 2)
    fld.Properties.Append fld.CreateProperty("ColumnWidths", dbText, "0cm;5cm")
    fld.Properties.Append fld.CreateProperty("LimitToList", dbBoolean, True)
End Sub