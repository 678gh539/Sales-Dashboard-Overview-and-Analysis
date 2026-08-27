# POWERBI Sales-Dashboard-Overview-and-Analysis
A dashboard that provides an overview of sales performance, including total sales, profit, sales trends, product categories, suppliers, brands, and monthly performance.

Total Items Sold	Total quantity of products sold	How many individual units customers purchased
Total Sales	Total sales/revenue from products	How much sales the business generated
Total Unit Price	Sum of product prices	Total of the listed unit prices; usually less useful as a KPI
Total Customers	Number of customers	Size of your customer base
Total Amount	Total amount recorded in payments	How much money was recorded as paid
Average Payment	Average payment amount	Typical amount paid per payment transaction
Total Products	Number of products in the product catalog	Size of your product inventory/catalog
Total Suppliers	Number of suppliers	How many suppliers provide products

# Excel Sales Dashboard

=XLOOKUP(B2,customers!$A$1:$A$53,customers!$B$1:$B$53,,)  XLOOKUP the first_name from the Customers table to the Orders table.
=XLOOKUP(B2,customers!$A$1:$A$53,customers!$D$1:$D$53,,) XLOOKUP THE email from customers to The orders table      

-- index match from the payments header to the orders table
=INDEX(products!A1:F51,MATCH(orders!$A2,products!$A$1:$A$51,0),MATCH(orders!I$1,products!$A$1:$F$1,0))

-- and mutiply the quatity * price to get the total sale
