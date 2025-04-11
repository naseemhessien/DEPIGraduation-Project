Create database DB_DownTime

use DB_DownTime

EXEC sp_configure 'Show Advanced Options', 1;
RECONFIGURE;
EXEC sp_configure 'Ad Hoc Distributed Queries', 1;
RECONFIGURE;
EXEC sp_configure 'Ole Automation Procedures', 1;
RECONFIGURE;
EXEC sp_configure 'xp_cmdshell', 1;
RECONFIGURE;

-- Import Data from excel sheets
SELECT * 
INTO line_productivity
FROM OPENROWSET(
    'Microsoft.ACE.OLEDB.12.0',
    'Excel 12.0;Database=E:\final\Manufacturing_Line_Productivity.xlsx;HDR=YES',
    'SELECT * FROM [Line_productivity$]'
);
select * from line_productivity;

SELECT * 
INTO Products
FROM OPENROWSET(
    'Microsoft.ACE.OLEDB.12.0',
    'Excel 12.0;Database=E:\final\Manufacturing_Line_Productivity.xlsx;HDR=YES',
    'SELECT * FROM [Products$]'
);
select * from Products;

SELECT * 
INTO Downtime_factors
FROM OPENROWSET(
    'Microsoft.ACE.OLEDB.12.0',
    'Excel 12.0;Database=E:\final\Manufacturing_Line_Productivity.xlsx;HDR=YES',
    'SELECT * FROM [Downtime_factors$]'
);
select * from Downtime_factors;

SELECT * 
INTO Line_downtime
FROM OPENROWSET(
    'Microsoft.ACE.OLEDB.12.0',
    'Excel 12.0;Database=E:\final\Manufacturing_Line_Productivity.xlsx;HDR=YES',
    'SELECT * FROM [Line_downtime$]'
);

select * from Line_productivity
select * from Line_downtime
select * from Normalized_line_Downtime
select * from Downtime_factors
select * from Products

select * from Line_downtime;
-- ÊäÙíÝ æÊÍæíá ÇáÈíÇäÇÊ Ýí ÌÏæá Products
EXEC sp_rename 'Products.Min batch time', 'Min_batch_time', 'COLUMN';

--ÊäÙíÝ æÊÍæíá ÇáÈíÇäÇÊ Ýí ÌÏæá Downtime_factors
EXEC sp_rename 'Downtime_factors.Operator Error', 'Operator_Error', 'COLUMN';

ALTER TABLE Downtime_factors
ALTER COLUMN Factor INT;


-- ÊäÙíÝ æÊÍæíá ÇáÈíÇäÇÊ Ýí ÌÏæá Line productivity
EXEC sp_rename 'Line_productivity.Start Time', 'Start_Time', 'COLUMN';
EXEC sp_rename 'Line_productivity.End Time', 'End_Time', 'COLUMN';

ALTER TABLE line_productivity  
DROP COLUMN  F7;

ALTER TABLE line_productivity
ALTER COLUMN Start_Time DATETIME; ----make change by heba make datetime instead of time

ALTER TABLE line_productivity
ALTER COLUMN End_Time DATETIME; ----make change by heba make datetime instead of time

ALTER TABLE line_productivity
ALTER COLUMN Date DATE;

ALTER TABLE line_productivity
ALTER COLUMN Batch INT;

-- ÊäÙíÝ ÈíÇäÇÊ Line_downtime
select * from Line_downtime;

CREATE TABLE Normalized_Line_Downtime (
    Batch INT,
    Downtime_factor NVARCHAR(50), 
    Downtime_mins INT
);

INSERT INTO Normalized_Line_Downtime (Batch, Downtime_factor, Downtime_mins)
SELECT 
    F1 AS Batch, 
    CAST(SUBSTRING(ColumnName, 2, LEN(ColumnName)-1) AS INT) - 1 AS Downtime_factor, 
    Value AS Downtime_mins
FROM Line_downtime
UNPIVOT (
    Value FOR ColumnName IN (F3, F4, F5, F6, F7, F8, F9, F10, F11, F12, F13)
) AS Unpvt
WHERE F1 IS NOT NULL AND Value IS NOT NULL;

-- ÅÏÎÇá "No Error" áÃí Batch áíÓ áå ÈíÇäÇÊ downtime
INSERT INTO Normalized_Line_Downtime (Batch, Downtime_factor, Downtime_mins)
SELECT 
    DISTINCT F1 AS Batch, 
    0 AS Downtime_factor, --make change by heba (change 'No Error' to 0)
    0 AS Downtime_mins
FROM Line_downtime
WHERE F1 NOT IN (SELECT DISTINCT Batch FROM Normalized_Line_Downtime);

SELECT * FROM Normalized_Line_Downtime ORDER BY Batch, Downtime_factor;

-- ÇáÈÍË Úä ÇáÞíã NULL Ýí ßá ÌÏæá
SELECT * FROM Line_productivity
WHERE Date IS NULL OR Product IS NULL OR Batch IS NULL OR Operator IS NULL OR Start_Time IS NULL OR End_Time IS NULL;

SELECT * FROM Normalized_Line_Downtime
WHERE Batch IS NULL OR Downtime_factor IS NULL OR Downtime_mins IS NULL;

SELECT * FROM Products
WHERE Product IS NULL OR Flavor IS NULL OR Size IS NULL OR  Min_batch_time IS NULL ;

SELECT * FROM Downtime_factors
WHERE Factor IS NULL OR Description IS NULL OR Operator_Error IS NULL ;

--ÇáÈÍË Úä ÇáÊßÑÇÑ Ýí ßá ÇáÌÏÇæá
--Normalized_Line_Downtime Table
WITH CTE1 AS (
    SELECT *, ROW_NUMBER() OVER (PARTITION BY Batch,Downtime_factor,Downtime_mins ORDER BY Downtime_mins) AS RowNum
    FROM Normalized_Line_Downtime
)
SELECT * FROM CTE1 WHERE RowNum > 1; --"No duplicate records found."

--Line_productivity Table
WITH CTE2 AS (
    SELECT *, ROW_NUMBER() OVER (PARTITION BY Date,Product,Batch,Operator,Start_Time,End_Time ORDER BY Batch) AS RowNum
    FROM Line_productivity
)
SELECT * FROM CTE2 WHERE RowNum > 1; --"No duplicate records found."

--Products Table
WITH CTE3 AS (
    SELECT *, ROW_NUMBER() OVER (PARTITION BY Product,Flavor,Size,Min_batch_time ORDER BY Product) AS RowNum
    FROM Products
)
SELECT * FROM CTE3 WHERE RowNum > 1; --"No duplicate records found."

--Downtime_factors Table
WITH CTE4 AS (
    SELECT *, ROW_NUMBER() OVER (PARTITION BY Factor,Description,Operator_Error ORDER BY Factor) AS RowNum
    FROM Downtime_factors
)
SELECT * FROM CTE4 WHERE RowNum > 1; --"No duplicate records found."





--------------------------- start make fact and dimension tables -----------
CREATE TABLE Dim_Product (
    Product_ID INT IDENTITY(1,1) PRIMARY KEY ,
    Product_Code VARCHAR(50) UNIQUE NOT NULL,
    Flavor VARCHAR(50),
    Size VARCHAR(20),
    Min_Batch_Time INT
);
---Insert Data from Products Table
INSERT INTO Dim_Product (Product_Code, Flavor, Size, Min_Batch_Time)
SELECT DISTINCT Product, Flavor, Size, Min_Batch_Time FROM Products;


CREATE TABLE Dim_Operator (
    Operator_ID INT IDENTITY(1,1) PRIMARY KEY,
    Operator_Name VARCHAR(50) UNIQUE NOT NULL
);
-- Insert Unique Operators
INSERT INTO Dim_Operator (Operator_Name)
SELECT DISTINCT Operator FROM Line_productivity;

CREATE TABLE Dim_Time (
    Time_ID INT IDENTITY(1,1) PRIMARY KEY ,
    Full_Date DATE,
    Day INT,
    Month INT,
    Year INT
);

--Insert Date Data
INSERT INTO Dim_Time (Full_Date, Day, Month, Year)
SELECT DISTINCT Date, DAY(Date), MONTH(Date), YEAR(Date) FROM Line_productivity;

CREATE TABLE Dim_Downtime (
    Downtime_ID INT IDENTITY(1,1) PRIMARY KEY ,
    Factor INT UNIQUE NOT NULL,
    Description VARCHAR(100),
    Operator_Error bit
);
--Insert Data from Downtime_factors
INSERT INTO Dim_Downtime (Factor, Description, Operator_Error)
SELECT Factor, Description, CASE WHEN Operator_Error = 'Yes' THEN 1 ELSE 0 END 
FROM Downtime_factors;



CREATE TABLE Fact_Production (
    Fact_ID INT IDENTITY(1,1) PRIMARY KEY ,
    Date_ID INT,
    Product_ID INT,
    Operator_ID INT,
    Batch INT,
    Start_Time DATETIME,
    End_Time DATETIME,
    Downtime_Duration INT DEFAULT 0,
    FOREIGN KEY (Date_ID) REFERENCES Dim_Time(Time_ID),
    FOREIGN KEY (Product_ID) REFERENCES Dim_Product(Product_ID),
    FOREIGN KEY (Operator_ID) REFERENCES Dim_Operator(Operator_ID)  
);

INSERT INTO Fact_Production (Date_ID, Product_ID, Operator_ID, Batch, Start_Time, End_Time, Downtime_Duration)
SELECT 
    T.Time_ID, 
    P.Product_ID, 
    O.Operator_ID, 
    L.Batch, 
    L.Start_Time, 
    L.End_Time, 
    SUM(UNPIVOTED.Downtime_mins) AS Downtime_Duration
FROM Line_productivity L
JOIN Dim_Time T ON L.Date = T.Full_Date
JOIN Dim_Product P ON L.Product = P.Product_Code
JOIN Dim_Operator O ON L.Operator = O.Operator_Name
LEFT JOIN Normalized_Line_Downtime UNPIVOTED ON L.Batch = UNPIVOTED.Batch
LEFT JOIN Dim_Downtime D ON UNPIVOTED.Downtime_factor = D.Factor 
GROUP BY T.Time_ID, 
         P.Product_ID, 
         O.Operator_ID, 
	     L.Batch,
		 L.Start_Time, 
         L.End_Time
order by L.Batch ;


select * from Fact_Production

-------------------------- ANALYSIS ----------------------

--Total Production Time per Product
SELECT DP.Product_Code, 
sum(CASE 
        WHEN DATEDIFF(MINUTE, CAST(FP.Start_Time AS DATETIME), CAST(FP.End_Time AS DATETIME)) < 0 
        THEN DATEDIFF(MINUTE, CAST(FP.Start_Time AS DATETIME), DATEADD(DAY, 1, CAST(FP.End_Time AS DATETIME)))
        ELSE DATEDIFF(MINUTE, CAST(FP.Start_Time AS DATETIME), CAST(FP.End_Time AS DATETIME))
    END) AS Total_Minutes
FROM Fact_Production FP
JOIN Dim_Product DP ON FP.Product_ID = DP.Product_ID
GROUP BY DP.Product_Code;


SELECT FP.Batch, DP.Product_Code, 
CASE 
        WHEN DATEDIFF(MINUTE, CAST(FP.Start_Time AS DATETIME), CAST(FP.End_Time AS DATETIME)) < 0 
        THEN DATEDIFF(MINUTE, CAST(FP.Start_Time AS DATETIME), DATEADD(DAY, 1, CAST(FP.End_Time AS DATETIME)))
        ELSE DATEDIFF(MINUTE, CAST(FP.Start_Time AS DATETIME), CAST(FP.End_Time AS DATETIME))
    END AS Total_Minutes
FROM Fact_Production FP
JOIN Dim_Product DP ON FP.Product_ID = DP.Product_ID
group by  FP.Batch, DP.Product_Code ,FP.Start_Time, FP.End_Time

select PRODUCT,
sum(CASE 
        WHEN DATEDIFF(MINUTE, FP.Start_Time, FP.End_Time) < 0 
        THEN DATEDIFF(MINUTE, FP.Start_Time, DATEADD(DAY, 1, FP.End_Time))
        ELSE DATEDIFF(MINUTE, FP.Start_Time, FP.End_Time)
    END) AS Total_Minutes
from line_productivity FP
group by PRODUCT


------------------------ Export Data ----------------------------------
-- ÊÕÏíÑ ÌÏæá Dim_Downtime Åáì Excel
INSERT INTO OPENROWSET('Microsoft.ACE.OLEDB.12.0',
    'Excel 12.0;Database=E:\final\cleanDataSQL.xlsx;',
    'SELECT * FROM [Dim_Downtime$]')
SELECT * FROM Dim_Downtime;

-- ÊÕÏíÑ ÌÏæá Dim_Operator Åáì Excel
INSERT INTO OPENROWSET('Microsoft.ACE.OLEDB.12.0',
    'Excel 12.0;Database=E:\final\cleanDataSQL.xlsx;',
    'SELECT * FROM [Dim_Operator$]')
SELECT * FROM Dim_Operator;

-- ÊÕÏíÑ ÌÏæá Dim_Product Åáì Excel
INSERT INTO OPENROWSET('Microsoft.ACE.OLEDB.12.0',
    'Excel 12.0;Database=E:\final\cleanDataSQL.xlsx;',
    'SELECT * FROM [Dim_Product$]')
SELECT * FROM Dim_Product;

-- ÊÕÏíÑ ÌÏæá Dim_Time Åáì Excel
INSERT INTO OPENROWSET('Microsoft.ACE.OLEDB.12.0',
    'Excel 12.0;Database=E:\final\cleanDataSQL.xlsx;',
    'SELECT * FROM [Dim_Time$]')
SELECT * FROM Dim_Time;

-- ÊÕÏíÑ ÌÏæá Fact_Production Åáì Excel
INSERT INTO OPENROWSET('Microsoft.ACE.OLEDB.12.0',
    'Excel 12.0;Database=E:\final\cleanDataSQL.xlsx;',
    'SELECT * FROM [Fact_Production$]')
SELECT * FROM Fact_Production;


-- ÊÕÏíÑ ÌÏæá Normalized_Line_Downtime Åáì Excel
INSERT INTO OPENROWSET('Microsoft.ACE.OLEDB.12.0',
    'Excel 12.0;Database=E:\final\cleanDataSQL.xlsx;',
    'SELECT * FROM [Normalized_Line_Downtime$]')
SELECT * FROM Normalized_Line_Downtime;




