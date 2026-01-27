-- ########### (Meta Database Engineer) MySQL 과목 내용 정리 

----------- Basics : Create, Replace Into, Alter, View, Update, Delete, 

-- 1) CREATE, INSERT INTO 

CREATE DATABASE IF NOT EXISTS Little_Lemon ; 
USE Little_Lemon ; 

CREATE TABLE Customers (
CustomerID INT NOT NULL PRIMARY KEY, 
FullName VARCHAR(100) NOT NULL, 
PhoneNumnber INT NOT NULL UNIQUE 
) ; 

CREATE DATABASE IF NOT EXISTS Lucky_Shrub ; 
USE Lucky_Shrub ; 

CREATE TABLE Orders (
OrderID INT, 
ClientID VARCHAR(10), 
ProductID VARCHAR(10), 
Quantity INT,  
Cost Decimal(6,2)) ; 

INSERT INTO Orders (
OrderID, ClientID, ProductID, Quantity, Cost) 
VALUES (1, "ABC". "P1", 10, 500), (2, "BCD". "P2", 20, 100), ... ; 

-- VIEW
CREATE VIEW BookingView AS  
SELECT BookingID, BookingDate, NumofGuests 
FROM Bookings 
WHERE 1=1 
AND NumofGuests > 3 
AND BookingDate < '20250101' ; 

-- 2) REPLACE INTO 
-- Replace 하려는 행이 없으면 Insert, 있으면 Replace 동작 
-- Kabasa의 cost를 replace from 17 to 20 
REPLACE INTO Courses (CourseName, Cost) VALUES ("Kabasa", 20) 

-- 3) ALTER
-- Columns관련(컬럼추가/삭제, 컬럼이름/타입변경, Constraints(PK/FK추가, UNIQUE추가 등), Index, Null/Not Null, Default 등 변경 
ALTER TABLE Courses ADD COLUMN Ingredients VARCHAR(255) ; 

-- 4) UPDATE  
UPDATE student_tbl 
SET college_address = 'Harper Bldg', home_address = 'xyz' 
WHERE department = 'engineering' 

-- 5) DELETE FROM 
DELETE FROM student_tbl 
WHERE department = 'engineering' ; 

----------- PROCEDURE 
-- 일련의 slq 문을 저장해두고, 필요할 때마다 호출하여 반복 실행할 수 있는 구조 
-- > 반복되는 sql 작업을 자동화하거나 재사용 가능하게 만드는 데 사용 

-- case 1) without parameter 
CREATE  PROCEDURE GetIalien() 
SELECT * FROM Customers WHERE Country = 'Italy'  ; 
-- call 
CALL GetItalien() ; 

-- case 2) with parameter (1) 
CREATE PROCEDURE GetProductBasedOnPrice (InputPrice INT) 
SELECT * FROM Products WHERE Price <= InputPrice ; 
-- call 
CALL GetProductBasedOnPrice (50) ; 

-- case 2) with parameter (2) 
CREATE PROCEDURE GetOrdersInRange (MinVal DECIMAL, MaxVal DECIMAL) 
SELECT * FROM Orders WHERE Cost >= MinVal AND Cost <= MaxVal  ; 
-- call 
CALL GetOrdersInRange (150,300) ; 

-- PROCEDURE 변경 필요시 DROP (ALTER 사용불가, DROP 후 CREATE) 
DROP PROCEDURE GetProductBasedOnPrice ; 

----------- STORED PROCEDURE 
-- example 1 (simple) 
DELIMITER // 
CREATE PROCEDURE GetAllClients() 
BEGIN 
	SELECT * FROM Clients ; -- when sql is complicated, use BEGIN ~ END 
END // 
DELIMETER ; -- to restore the terminator 
-- call 
CALL GetAllClients() ; 

-- example 2
DELIMITER // 
CREATE Procedure GetDiscount(OrderIDInput INT) ; 

BEGIN
	DECLARE cost_after_discount DECIMAL(7,2) ; 
	DECLARE current_cost DECIMAL(7,2) ; 
	DECLARE order_quantity INT ; 
	
	SELECT Quantity INTO order_quantity FROM Orders WHERE OrderID = OrderIDInput ; 
	SELECT Cost INTO current_cost FROM Orders WHERE OrderID = OrderIDInput ; 
	
	IF order_quantity >= 20 THEN 
		SET cost_after_discount = current_cost - (current_cost * 0.2) ;  -- 20% discount
	ELSEIF order_quantity >= 10 THEN 
		SET cost_after_discount = current_cost - (current_cost * 0.1) ;  -- 10% discount
	ELSE 
		SET cost_after_discount = current_cost ; -- no discount 
	END IF 
SELECT cost_after_discount 

END // 
DELIMITER ; 

----------- FUNCTION 

-- PROCEDURE는 데이터조작, 처리, 실행 등 목적을 가진 반면, FUNCTION은 계산 후 단일결과반환 특화 
-- PROCEDURE가 CALL 호출하여 단독실행하는 반면, FUNCTION은 SELECT 함수명 또는 또다른 SQL문 내 호출 
-- PROCEDURE은 RETURN 값이 없는 (대신 OUT또는INOUT) 반면, FUNCTION은 RETURN 1개 반환함. 

-- 3단구성 
-- (1) CREATE FUNCTION : 함수이름, 입력파라미터 정의
-- (2) RETURNS DECIMAL(5,2) DETERMINISTIC : 반환타입 정의 
-- (3) RETURN (SELECT ~~~ ) ; : 반환값 지정 

SELECT AVG(Cost) FROM Orders ; 

DELIMITER // 
CREATE FUNCTION GetCostAvg() RETURNS DECIMAL(5,2) DETERMINISTIC 
BEGIN 
	RETURN (SELECT AVG(Cost) From Orders) ; 
END // 
DELIMITER ; 

SELECT GetCostAvg() ; 

----------- TRIGGER 
-- TRIGGER은 테이블에 INSERT, UPDATE, DELETE 등의 작업이 발생할 때 자동으로 실행하는 SQL 블록
-- 데이터무결성 유지(자동 검증), 로그자동기록, 자동계산, 감사(audit)목적 등을 위해 사용 
-- TRIGGER 종류 : BEFORE/AFTER INSERT, BEFORE/AFTER UPDATE, BEFORE/AFTER DELETE

-- solution 1 : AFTER INSERT ON Products, INSERT INTO Notifications 

DELIMITER //

CREATE TRIGGER ProductSellPriceInsertCheck 
    AFTER INSERT  
    ON Products FOR EACH ROW  
    BEGIN
    IF NEW.SellPrice <= NEW.BuyPrice THEN
        INSERT INTO Notifications(Notification,DateTime) 
        VALUES(CONCAT('A SellPrice same or less than the BuyPrice was inserted for ProductID ', NEW.ProductID), NOW()); 
    END IF;
    END //

DELIMITER ;

-- solution 2 : AFTER UPDATE ON Products, INSERT INTO Notifications 

DELIMITER //

CREATE TRIGGER ProductSellPriceUpdateCheck 
    AFTER UPDATE  
    ON Products FOR EACH ROW  
    BEGIN
    IF NEW.SellPrice <= NEW.BuyPrice THEN
        INSERT INTO Notifications(Notification,DateTime) 
        VALUES(CONCAT(NEW.ProductID,' was updated with a SellPrice of ', NEW.SellPrice,' which is the same or less than the BuyPrice'), NOW()); 
    END IF;
    END //

DELIMITER ;

-- solution 3 : AFTER DELETE ON Products, INSERT INTO Notifications 

DELIMITER //

CREATE TRIGGER NotifyProductDelete 
    AFTER DELETE   
    ON Products FOR EACH ROW   
    INSERT INTO Notifications(Notification, DateTime) 
    VALUES(CONCAT('The product with a ProductID ', OLD.ProductID,' was deleted'), NOW()); 
END //
DELIMITER ;

----------- Advanced MySQL Topics   
-- Task 1 solution:
-- Lucky Shrub need to find out what their average cost was for a product in 2022.

CREATE FUNCTION FindAverageCost(YearInput INT) 
RETURNS DECIMAL(10,2) DETERMINISTIC 
RETURN (SELECT AVG(Cost) FROM Orders WHERE YEAR(Date) = YearInput); 

-- Task 2 solution:
-- Lucky Shrub need to evaluate the sales patterns for bags of artificial grass over the last three years. Help them out using the following steps:
DELIMITER // 
CREATE PROCEDURE EvaluateProduct(IN product_id VARCHAR(10), OUT SoldItemsIn2020 INT, OUT SoldItemsIn2021 INT, OUT SoldItemsIn2022 INT)
BEGIN
SELECT SUM(Quantity) INTO SoldItemsIn2020 FROM Orders WHERE ProductID=product_id AND YEAR(Date)=2020; 
SELECT SUM(Quantity) INTO SoldItemsIn2021 FROM Orders WHERE ProductID=product_id AND YEAR(Date)=2021;
SELECT SUM(Quantity) INTO SoldItemsIn2022 FROM Orders WHERE ProductID=product_id AND YEAR(Date)=2022; 
END //
DELIMITER ;

CALL EvaluateProduct('P1', @sold_items_2020, @sold_items_2021, @sold_items_2022);

-- Task 3 Solution
/*
 Lucky Shrub need to automate the orders process in their database. 
The database must insert a new record of data in response to the insertion of a new order in the Orders table. 
This new record of data must contain a new ID and the current date and time.
*/

CREATE TRIGGER UpdateAudit AFTER INSERT 
ON Orders 
FOR EACH ROW 
INSERT INTO Audit (OrderDateTime) 
VALUES (Current_timestamp);

-- Task 4 Solution
/*
Lucky Shrub need location data for their clients and employees. 
To help them out, create an optimized query that outputs the following data:
*/

SELECT Employees.FullName, Addresses.Street, Addresses.County 
FROM Employees INNER JOIN Addresses 
ON Employees.AddressID = Addresses.AddressID
UNION
SELECT Clients.FullName, Addresses.Street, Addresses.County 
FROM Clients INNER JOIN Addresses ON Clients.AddressID = Addresses.AddressID 
ORDER BY Street;

-- Task 5 Solution
/*
Lucky Shrub need to find out what quantities of wood panels they are selling. 
The wood panels product has a Product ID of P2. 
The following query returns the total quantity of this product as sold in the years 2020, 2021 and 2022:
*/

SELECT CONCAT(SUM(Cost), "(2020)") AS "Total sum of P2 Product" 
FROM Orders WHERE YEAR(Date) = 2020 AND ProductID = "P2" UNION 
SELECT CONCAT(SUM(Cost), "(2021)") FROM Orders 
WHERE YEAR(Date) = 2021 AND ProductID = "P2"
UNION SELECT CONCAT(SUM(Cost), "(2022)") FROM Orders 
WHERE YEAR(Date) = 2022 AND ProductID = "P2";

-- You are tasked to optimize this query by recreating it as a common table expression (CTE). 

WITH
P2_Sales_2020 AS (SELECT CONCAT(SUM(Cost), " (2020)") AS "Total sum of P2 Product" FROM Orders WHERE YEAR(Date) = 2020 AND ProductID= "P2"),
P2_Sales_2021 AS (SELECT CONCAT(SUM(Cost), " (2021)") AS "Total sum of P2 Product" FROM Orders WHERE YEAR(Date) = 2021 AND ProductID= "P2"),
P2_Sales_2022 AS (SELECT CONCAT(SUM(Cost), " (2022)") AS "Total sum of P2 Product" FROM Orders WHERE YEAR(Date) = 2022 AND ProductID= "P2")
SELECT * FROM P2_Sales_2020
UNION
SELECT * FROM P2_Sales_2021
UNION
SELECT * FROM P2_Sales_2022;

-- Task 6 Solution
/*
Lucky Shrub want to know more about the activities of the clients who use their online store. 
The system logs the ClientID and the ProductID information for each activity in a JSON Properties column inside the Activity table.
This occurs while clients browse through Lucky Shrub products online.
The following screenshot shows the Activity table.
*/

SELECT Activity.Properties ->>'$.ClientID' 
AS ClientID, Activity.Properties ->>'$.ProductID' 
AS ProductID, Clients.FullName, Clients.ContactNumber 
FROM Clients RIGHT JOIN Activity 
ON Clients.ClientID = Activity.Properties ->>'$.ClientID';

-- Task 7 Solution
/*
Lucky Shrub need to find out how much revenue their top selling product generated. 
Create a stored procedure called GetProfit that returns the overall profits generated by a specific product in a specific year. 
This should be based on the user input of the ProductID and Year.
For example, the output result of GetProfit() procedure with the P1 ProductID and Year 2020 is displayed in the screenshot below.  
*/

DELIMITER //
CREATE PROCEDURE GetProfit(IN product_id VARCHAR(10), IN YearInput INT)
BEGIN
DECLARE profit DEC(7,2) DEFAULT 0.0; 
DECLARE sold_quantity, buy_price, sell_price INT DEFAULT 0;
SELECT SUM(Quantity) INTO sold_quantity FROM Orders WHERE ProductID = product_id AND YEAR(Date) = YearInput; 
SELECT BuyPrice INTO buy_price FROM Products WHERE ProductID = product_id; 
SELECT SellPrice INTO sell_price FROM Products WHERE ProductID = product_id;
SET profit = (sell_price * sold_quantity) - (buy_price * sold_quantity);
Select profit; 
END //
DELIMITER ;

-- Task 8 solution
/*
Lucky Shrub need a summary of their client's details, including their addresses, order details and the products they purchased. 
Help them out by creating a virtual table called DataSummary that joins together the four tables that contain this data. 
*/

CREATE VIEW DataSummary AS 
SELECT Clients.FullName, Clients.ContactNumber, Addresses.County, Products.ProductName, Orders.ProductID, Orders.Cost, Orders.Date 
FROM Clients 
INNER JOIN Addresses ON Clients.AddressID = Addresses.AddressID 
INNER JOIN Orders ON Clients.ClientID = Orders.ClientID INNER JOIN Products ON Orders.ProductID = Products.ProductID 
WHERE YEAR(Orders.Date) = 2022 
ORDER BY Orders.Cost DESC ; 
