USE [master]
GO
/****** Object:  Database [CartaoDeCredito]    Script Date: 11/10/2024 11:29:42 ******/
CREATE DATABASE [CartaoDeCredito]
 CONTAINMENT = NONE
 ON  PRIMARY 
( NAME = N'CartaoDeCredito', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL16.SQLEXPRESS\MSSQL\DATA\CartaoDeCredito.mdf' , SIZE = 8192KB , MAXSIZE = UNLIMITED, FILEGROWTH = 65536KB )
 LOG ON 
( NAME = N'CartaoDeCredito_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL16.SQLEXPRESS\MSSQL\DATA\CartaoDeCredito_log.ldf' , SIZE = 8192KB , MAXSIZE = 2048GB , FILEGROWTH = 65536KB )
 WITH CATALOG_COLLATION = DATABASE_DEFAULT, LEDGER = OFF
GO
ALTER DATABASE [CartaoDeCredito] SET COMPATIBILITY_LEVEL = 160
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [CartaoDeCredito].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO
ALTER DATABASE [CartaoDeCredito] SET ANSI_NULL_DEFAULT OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET ANSI_NULLS OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET ANSI_PADDING OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET ANSI_WARNINGS OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET ARITHABORT OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET AUTO_CLOSE OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET AUTO_SHRINK OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET AUTO_UPDATE_STATISTICS ON 
GO
ALTER DATABASE [CartaoDeCredito] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET CURSOR_DEFAULT  GLOBAL 
GO
ALTER DATABASE [CartaoDeCredito] SET CONCAT_NULL_YIELDS_NULL OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET NUMERIC_ROUNDABORT OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET QUOTED_IDENTIFIER OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET RECURSIVE_TRIGGERS OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET  DISABLE_BROKER 
GO
ALTER DATABASE [CartaoDeCredito] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET TRUSTWORTHY OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET PARAMETERIZATION SIMPLE 
GO
ALTER DATABASE [CartaoDeCredito] SET READ_COMMITTED_SNAPSHOT OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET HONOR_BROKER_PRIORITY OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET RECOVERY SIMPLE 
GO
ALTER DATABASE [CartaoDeCredito] SET  MULTI_USER 
GO
ALTER DATABASE [CartaoDeCredito] SET PAGE_VERIFY CHECKSUM  
GO
ALTER DATABASE [CartaoDeCredito] SET DB_CHAINING OFF 
GO
ALTER DATABASE [CartaoDeCredito] SET FILESTREAM( NON_TRANSACTED_ACCESS = OFF ) 
GO
ALTER DATABASE [CartaoDeCredito] SET TARGET_RECOVERY_TIME = 60 SECONDS 
GO
ALTER DATABASE [CartaoDeCredito] SET DELAYED_DURABILITY = DISABLED 
GO
ALTER DATABASE [CartaoDeCredito] SET ACCELERATED_DATABASE_RECOVERY = OFF  
GO
ALTER DATABASE [CartaoDeCredito] SET QUERY_STORE = ON
GO
ALTER DATABASE [CartaoDeCredito] SET QUERY_STORE (OPERATION_MODE = READ_WRITE, CLEANUP_POLICY = (STALE_QUERY_THRESHOLD_DAYS = 30), DATA_FLUSH_INTERVAL_SECONDS = 900, INTERVAL_LENGTH_MINUTES = 60, MAX_STORAGE_SIZE_MB = 1000, QUERY_CAPTURE_MODE = AUTO, SIZE_BASED_CLEANUP_MODE = AUTO, MAX_PLANS_PER_QUERY = 200, WAIT_STATS_CAPTURE_MODE = ON)
GO
USE [CartaoDeCredito]
GO
/****** Object:  Table [dbo].[tblClientes]    Script Date: 11/10/2024 11:29:42 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblClientes](
	[idCliente] [int] IDENTITY(1,1) NOT NULL,
	[nome] [varchar](100) NULL,
	[nuCartao] [varchar](20) NULL,
 CONSTRAINT [PK_tblClientes] PRIMARY KEY CLUSTERED 
(
	[idCliente] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[tblTransacao]    Script Date: 11/10/2024 11:29:43 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblTransacao](
	[idTransacao] [int] IDENTITY(1,1) NOT NULL,
	[nuCartao] [varchar](20) NULL,
	[valorTransacao] [decimal](10, 2) NULL,
	[dataTransacao] [date] NULL,
	[descricao] [varchar](200) NULL,
 CONSTRAINT [PK_tblTransacoes] PRIMARY KEY CLUSTERED 
(
	[idTransacao] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
USE [master]
GO
ALTER DATABASE [CartaoDeCredito] SET  READ_WRITE 
GO
