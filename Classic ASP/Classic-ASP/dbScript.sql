/*Table Script*/

USE [Work]
GO

/****** Object:  Table [dbo].[emp]    Script Date: 5/25/2022 8:36:15 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[emp](
	[empno] [int] IDENTITY(1,1) NOT NULL,
	[ename] [varchar](10) NULL,
	[job] [varchar](100) NULL,
	[mgr] [int] NULL,
	[hiredate] [datetime] NULL,
	[sal] [numeric](7, 2) NULL,
	[comm] [numeric](7, 2) NULL,
	[dept] [int] NULL,
 CONSTRAINT [PK__emp__AF4C318A81F586AB] PRIMARY KEY CLUSTERED 
(
	[empno] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO




USE [Work]
GO

/****** Object:  Table [dbo].[dept]    Script Date: 5/25/2022 8:36:26 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[dept](
	[deptno] [int] NOT NULL,
	[dname] [varchar](14) NULL,
	[loc] [varchar](13) NULL
) ON [PRIMARY]
GO



/************************/

/*Procedure Generation8*/





USE [Work]
GO
/****** Object:  StoredProcedure [dbo].[getempdata]    Script Date: 5/25/2022 8:35:09 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER proc [dbo].[getempdata]
as 
begin
set nocount on
select empno,ename,job,sal,d.dname,d.loc from emp e inner join dept d
on e.dept=d.deptno
end




USE [Work]
GO
/****** Object:  StoredProcedure [dbo].[saveemp]    Script Date: 5/25/2022 8:35:25 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER proc [dbo].[saveemp]
(
 @ename varchar(400),
 @job varchar(400),
 @salary numeric(7,2)
)
as
begin
insert into emp(ename, job, sal,dept) values( @ename,@job, @salary,2)
end




