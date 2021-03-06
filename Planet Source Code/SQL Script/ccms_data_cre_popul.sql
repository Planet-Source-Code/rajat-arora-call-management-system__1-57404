if exists (select * from sysobjects where id = object_id(N'[dbo].[call_registration]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[call_registration]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[customer_details]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[customer_details]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[department]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[department]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[employee_details]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[employee_details]
GO

CREATE TABLE [dbo].[call_registration] (
	[notification_no] [int] NULL ,
	[cust_no] [int] NULL ,
	[name] [nvarchar] (50) NULL ,
	[status] [nvarchar] (50) NULL ,
	[date_reg] [smalldatetime] NULL ,
	[date_close] [smalldatetime] NULL ,
	[zone] [nvarchar] (50) NULL ,
	[product] [nvarchar] (50) NULL ,
	[call_type] [nvarchar] (50) NULL ,
	[taker] [nvarchar] (50) NULL ,
	[service_request] [nvarchar] (50) NULL ,
	[priority] [nvarchar] (50) NULL ,
	[bd] [bit] NOT NULL ,
	[bd_text] [nvarchar] (50) NULL ,
	[remarks] [nvarchar] (50) NULL ,
	[sign_by] [nvarchar] (50) NULL ,
	[attend_by] [nvarchar] (50) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[customer_details] (
	[cust_no] [int] NULL ,
	[title] [int] NULL ,
	[fname] [nvarchar] (50) NULL ,
	[lname] [nvarchar] (50) NULL ,
	[houseno] [nvarchar] (50) NULL ,
	[add1] [nvarchar] (50) NULL ,
	[add2] [nvarchar] (50) NULL ,
	[locality] [nvarchar] (50) NULL ,
	[landmark] [nvarchar] (50) NULL ,
	[city] [nvarchar] (50) NULL ,
	[pincode] [int] NULL ,
	[phone] [nvarchar] (50) NULL ,
	[addproof] [nvarchar] (50) NULL ,
	[appliance] [nvarchar] (50) NULL ,
	[cust_type] [nvarchar] (50) NULL ,
	[dop] [smalldatetime] NULL ,
	[plant] [int] NULL ,
	[d_zone] [nvarchar] (50) NULL ,
	[sp_code] [nvarchar] (50) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[department] (
	[department] [nvarchar] (50) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[employee_details] (
	[emp_no] [int] NULL ,
	[fname] [nvarchar] (50) NULL ,
	[lname] [nvarchar] (50) NULL ,
	[add1] [nvarchar] (50) NULL ,
	[add2] [nvarchar] (50) NULL ,
	[city] [nvarchar] (50) NULL ,
	[pincode] [int] NULL ,
	[phone] [nvarchar] (50) NULL ,
	[department] [nvarchar] (50) NULL ,
	[designation] [nvarchar] (50) NULL ,
	[salary] [nvarchar] (50) NULL 
) ON [PRIMARY]
GO

insert into customer_details values(12346,1,"DHIRAJ","BALOTIA","44/16","ASHOK NAGAR","ASHOK NAGAR","TILAK NAGAR","NR SANATAN DHARAM MANDIR","NEW DELHI",110018,25149933,"VOTER ID CARD","AC 2 TON","Business","11/13/04",4510,"NORTH","H01")

insert into call_registration values(2342,12346,"DHIRAJ BALOTIA","COMPLETE","11/13/04","11/13/04","NORTH","250 LTS DD","Z1","RAJAT","NOT WORKING - URGENT - ASAP","H",1,"CUSTOMER VERY IRRETATED","NOT WORKING - URGENT - ASAP","ABC","MANOJ")

INSERT INTO EMPLOYEE_DETAILS VALUES(1101,"Rohit","Kumar","RZ 96","Vikas Puri","New Delhi","110018","25534349","Sales","Office Executive",10000)

INSERT INTO DEPARTMENT VALUES("SALES")

INSERT INTO DEPARTMENT VALUES("PURCHASE")

INSERT INTO DEPARTMENT VALUES("IT")

select * from customer_details

select * from call_registration

Select * from employee_details

select * from department

CREATE TABLE users
(
username varchar(30),
userid varchar(30)

)

INSERT INTO users values("dhiraj","dhiraj_hotline")
INSERT INTO users values("chintan arora","chintan_tiger")

select * from users