--CREATE DATABASE TestDB;
--DROP DATABASE TestDB;
--DROP TABLE TestDB.dbo.TestTable;
CREATE TABLE TestDB.dbo.TestTable (
 [id]  nvarchar(max) ,
 [first_name]  nvarchar(max) ,
 [last_name]  nvarchar(max) ,
 [email]  nvarchar(max) ,
 [gender]  nvarchar(max) ,
 [ip_address]  nvarchar(max) ,
);
SELECT * FROM TestDB.dbo.TestTable;