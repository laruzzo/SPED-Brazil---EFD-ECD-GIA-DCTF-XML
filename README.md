# SPED-Brazil---EFD-ICMS-IPI
This repository is part of the project SPED Brazil, specifically to generate the files related to EFD ICMS IPI. Code is originally in VBA and databse is MySQL, and I worked to convert it to Python. There is another repository for the SQL data structure

This project was totally developed by me, alone and was used in production environment to generate SPED files to my own company, Overture Brewery. So some part of the code are related to brewery operations, but not necessarily is only to be used on brewery. It is aligned to all business.

I am sharing this code as my brewery close the operations and now I would like to help other people to keep going with SPED development.

If you find this code usefull, I would be happy to be invited by you to contribute professionally to develop a commercial solution based on this.


VBA Code Analysis - Tax Recording System

Overview

The VBA code analyzed is a module called modGeraArquivo_EFD_ICMS_IPI, which is part of a tax recordkeeping system. Its main function is to generate EFD (Digital Tax Recording) files for ICMS and IPI.

Main Structure

1. Database Connection

•
Function: ConnectToDataBase()

•
Technology: MySQL via ODBC

•
Features:

•
Reads connection settings from the tbDBConn table

•
Uses MySQL ODBC 8.0 Unicode driver

•
Establishes a global connection using the Conn variable

2. Main Function

•
Function: Gerar_EFD_ICMS_IPI()

•
Parameters:

•
cDtIni: Start date

•
cDtFim: End date

•
clocal: Output directory

•
cDtINI_Contabil: Accounting start date

•
cIDIventario: Inventory ID

Identified Functions

1. EFD File Generation

•
Creates text file with a format specific to the Federal Revenue Service

•
File name: EFD_ICMS_IPI_[month]_[year].txt

•
Block structure (0, B, C, D, E)

2. Data Processing

•
Active Suppliers: Identifies suppliers with activity during the period

•
Active Products: Lists products with activity (sales, purchases, fixed assets)

•
Customers: Extracts customers with sales during the period

•
Transportation: Processes CTe (Electronic Bill of Lading)

•
Electric Energy: Processes energy invoices

•
SAT: Processes electronic tax receipts

3. EFD Blocks

Block 0 - Opening, Identification, and References

•
Company Data

•
Participant Registration (Customers/Suppliers)

•
Product Registration

•
Accounting Chart of Accounts

•
Cost Centers

Block B - ISS (from 2019)

•
ISS Bookkeeping and Calculation

Block C - Tax Documents I (ICMS/IPI)

•
Incoming and Outgoing Invoices

•
SAT Tax Receipts

•
Electricity Bills

Block D - Tax Documents II (ICMS Services)

•
Electronic Waybills (CTe)

Block E - ICMS and IPI Calculation

•
Monthly Calculation Calculations

•
Amounts to be Collected

Identified Database Tables

Main Entities

1.
tbCompany - Company Data

2.
tbAccountant - Accountant Data

3.
tbCustomer - Customer Registration

4.
tbSupplier - Supplier Registration

5.
tbProdCad - Product Registration

6.
tbSales / tbVendasDet - Sales and Items

7.
tbCompras / tbComprasDet - Purchases and Items

8.
tbTransportes - Bills of Lading

9.
tbImpedabilidades / tbImpedizadoCadastro - Fixed Assets

10.
tbVendasSAT / tbVendasSATDet - SAT Coupons

11.
tbEnergia - Electricity Bills

12.
tbResumo_ICMS - ICMS Tax Calculation Summary

13.
tbPlanoContasContabeis - Chart of Accounts

Temporary Tables

•
tbFornecedor_Ativo_temp - Active Suppliers in the Period

•
tbCadProd_Ativo_temp - Active Products in the Period

Technical Features

Versioning

•
Support for different EFD versions based on the date:

•
Until December 31, 2017: Version 011

•
From January 1, 2019: Version 015

•
From January 1, 2022: Version 016

•
From January 1, 2023: Version 017

Data Processing

•
Date formatting for different formats

•
Tax calculations (ICMS, IPI, PIS, COFINS)

•
Grouping and consolidation

•
Specific validations by transaction type

Complexity

•
Approximately 2,254 lines of code

•
Multiple complex SQL queries

•
Business logic specific to Brazilian tax legislation

•
Handling different types of tax documents

Conversion Challenges

1.
Access Dependencies: DAO, CurrentDb()

2.
MySQL Connection: Windows-specific ADODB

3.
File Formatting: EFD-specific Structure

4.
Tax Logic: Specific Knowledge of Legislation

5.
Performance: Multiple Queries and Large-Volume Processing

Next Steps

1.
Map All Tables and Relationships

2.
Create an Equivalent Data Structure in Python

3.
Implement Database Connection

4.
Translate Business Logic

5.
Implement EFD File Generation

6.
Create User Interface
