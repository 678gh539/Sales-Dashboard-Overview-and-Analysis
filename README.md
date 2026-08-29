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



# SQL Data Analysis Projects

-- i created table partition, primey key has to be connected orders has cusmoter_id 
-- and table customer has a customer_id

create table order_01(
order_id int not null,
order_date date not null,
customer_id int not null references customers(customer_id)
)partition by list (order_id);

insert into order_01
select a.order_id, a.order_date, a.customer_id
from orders a
left outer join customers b on a.customer_id = b.customer_id

-- 
create table customer_id_01
partition of order_01
for values in (22);
create table customer_id_02
partition of order_01
for values in (44);
create table customer_id_03
partition of order_01
for values in (11);
create table customer_id_default
partition of order_01
default;

insert into order_01
select order_id, order_date, customer_id
from orders;

-- must have group by 1
select order_date, count(*)
from customer_id_02
group by 1;

create table payment_partition(
payment_id int not null,
payment_date date not null,
amount decimal not null
)partition by list (payment_id);

-- always connect to your created partition
-- always put default 
create table payments_01
partition of payment_partition
for values in (1);
create table payment_default
partition of payment_partition
default;

insert into payment_partition
select
payment_id, payment_date, amount
from payments;

select payment_id, count(*)
from payments_01
group by 1;
select payment_date, count(*)
from payments_01
group by 1;
select payment_id, count(*)
from payment_default
group by 1;

-- temp table
create temporary table My_temp as 
select order_id, payment_date from payments
union all 
select order_id, cancel_date from orders;

select * from My_temp;



-- create view
create view payment_stats as select to_char(payment_date, 'yyyy-mm'), order_id
count(payment_id), sum(amount) as total amount
from payments
group by to_char(payment_date, 'yyyy-mm'), order_id
order by order_id, to_char(payment_date, 'yyyy-mm');

create table products_partition
(product_id int not null, product_name varchar(100), category_id int not null,
supplier_id int not null, price numeric(10, 2), stock int not null)
partition by hash (category_id, supplier_id);

create table products_partition_hash0
partition of products_partition
for values with (modulus 3, remainder 0);

create table products_partition_hash1
partition of products_partition
for values with (modulus 3, remainder 1);

create table products_partition_hash2
partition of products_partition
for values with (modulus 3, remainder 2);

insert into products_partition
select product_id, product_name, category_id, supplier_id, price, stock
from products;

select 'hash0', count(*)
from products_partition_hash0
union
select 'hash1', count(*)
from products_partition_hash1
union
select 'hash2', count(*)
from products_partition_hash2;

-- For the parent table, use id SERIAL if you want to use the default value.
CREATE TABLE orders_date (
    order_id INT PRIMARY KEY,
    customer_id INT,
    order_datetime TIMESTAMP NOT NULL
);
-- child table, make sure to create a triage_datetime column so you can see both tables when you select from the parent table.
create table orders_2026  
	order_status VARCHAR(50),
	triage_datetime timestamp
) INHERITS (orders_date);

insert into orders_2026 values 
(1, 1, '2026-08-16 10:30:00', 'delivered');

-- i remove the order status because we cannot update to child table from parent table
insert into orders_date values
(3, 3, '2026-08-15 10:30:00');
insert into orders_date values
(2, 2, '2026-08-15 10:30:00');

select * from orders_2026; 
select * from orders_date;
select * from only orders_date;

-- creating view

create view payment_stats as select to_char(payment_date, 'yyyy-mm-'), order_id,
count(payment_id), sum(amount) as total_amount
from payments
group by to_char(payment_date, 'yyyy-mm-dd'), order_id
order by order_id, to_char(payment_date, 'yyyy-mm--dd');

select * from payment_stats;

create view payment_stats_date as 
select order_id, payment_id, amount, payment_date
from payments
where payment_date = ('2026-07-01');

select * from payment_stats;
select * from payment_stats_date;
select * from only payment_stats_date;

-- you can also insert to view 
insert into payment_stats_date values
(51, 51, 50000.00, '2026-07-01');

create materialized view orders_stats as
select (order_id), (customer_id),
to_char(order_date, 'yyyy-mm-dd') as order_month, order_status 
from orders
group by 1, 2
order by 2, 1
with no data;

-- need to refresh
refresh materialized view orders_stats;
refresh materialized view concurrently orders_stats;
select * from orders_stats;

create recursive view v_fibonacci(a, b) as
select 1 as a, 1 as b
union all
select b, a+b
from v_fibonacci
where b < 200;

select * from v_fibonacci;

-- + 1 AS level, + 1 → increase it by 1
-- Repeatedly follow a relationship until there are no more rows to find.
create recursive view v_suppliers(supplier_name, phone, level) as
select supplier_name, phone, 0 as level
from suppliers
where phone like ('%__')
union all
select a.supplier_name, a.phone, b.level + 1 as level
from suppliers a
inner join v_suppliers b
on a.supplier_name = b.phone;

select * from v_suppliers;

-- In a view, you can join tables even if the two tables do not have a foreign key relationship.
create view customers_payment as
select a.customer_id, a.first_name, a.last_name, a.email, a.phone, a.address,
b.amount from customers a
left outer join payments b on a.customer_id = payment_id;

select * from customers_payment; 

-- Materialized view is like a saved result of a SQL query.
create materialized view C_O_P_stats as
select a.first_name, a.last_name
from customers a left outer join orders b on a.customer_id = b.customer_id
left outer join payments c on b.order_id = c.order_id
group by 1, 2
order by 1, 2
with no data;
refresh materialized view C_O_P_stats;
select * from C_O_P_stats;

create materialized view C_O_P_stats_1 as
select b.order_id, b.order_status, to_char(order_date, 'yyyy-mm-dd')
from customers a left outer join orders b on a.customer_id = b.customer_id
left outer join payments c on b.order_id = c.order_id
group by 1, 2, 3
order by 1, 3
with no data;

refresh materialized view C_O_P_stats_1;
select * from C_O_P_stats_1;

create materialized view C_O_P_stats_2 as
select a.first_name, a.last_name, b.order_id, b.order_status, to_char(order_date, 'yyyy-mm-dd')
from customers a left outer join orders b on a.customer_id = b.customer_id
left outer join payments c on b.order_id = c.order_id
group by 1, 2, 3, 4, 5
order by 1, 3
with no data;

refresh materialized view C_O_P_stats_2;
select * from C_O_P_stats_2;

select * from information_schema.routines
where routine_schema = 'payments'
and routine type = 'procedure'

-- create function and trigger
create function customers_trigger()
returns trigger
language plpgsql
as $$
begin
if new.last_name is null or new.first_name is null then
raise exception 'name cannot be null';
else new.first_name = trim(new.first_name);
     new.last_name = trim(new.last_name);
     new.full_name = concat(new.last_name, ', ', new.first_name);
	 return new;
	 end if;
end;
$$

create trigger customers_trigger_1
before insert
on customers
for each row execute procedure customers_trigger();

insert into customers values
(52, 'Jhonathan', 'Gonzaga', 'something');

select * from products;


-- array example {"Potato Chips"}
with resources as(
SELECT
    category_id,
    ARRAY_AGG(distinct product_name order by product_name) AS r_products
FROM products
GROUP BY category_id)
select r1.category_id, r1.r_products, r2.category_id
from resources r1
left outer join resources r2 on r1.category_id = r2.category_id
and r1.r_products = r2.r_products;

SELECT
    customer_id,
    ARRAY_AGG(order_id ORDER BY order_id) AS order_ids
FROM orders
GROUP BY customer_id;

with resources as(
SELECT
    customer_id,
    ARRAY_AGG(distinct order_id order by order_id) AS resources_array
FROM orders
GROUP BY customer_id)
select r1.customer_id, r1.resources_array, r2.customer_id
from resources r1
left outer join resources r2 on r1.customer_id = r2.customer_id
and r1.resources_array = r2.resources_array;

-- json and jason function
select '{"first_name": "Jhonathan", "last_name": "Gonzaga"}'::json->'first_name';

select jsonb_build_object (
'customer_id', customer_id,
'first_name', first_name,
'last_name', last_name
 ) as customers_json
 from customers;
 
--                                   JOIN
-- In the subquery, I use CURRENT_DATE to calculate the years_difference.
 
SELECT EXTRACT(YEAR FROM AGE(CURRENT_DATE, payment_date)) AS years_difference, order_date
from payments a left outer join orders b on a.order_id = b.order_id
where payment_date >= '2025-08-18';

-- inner join
select a.order_id, a.order_date, a.order_status, b.payment_date
from orders a inner join payments b on a.order_id = b.order_id;
-- inner join with USING is used with JOIN when both tables have a column with the same name.
select a.order_id, a.order_date, a.order_status, b.payment_date
from orders a inner join payments b using(order_id);

select count(distinct a.order_id)
from orders a inner join payments b using(order_id);

-- With a CROSS JOIN, you can combine two tables without a foreign key. 
select a.payment_method, b.price
from payments a cross join products b;

-- With a FULL JOIN, you can combine two tables without a foreign key.
-- You can't join the two tables without an integer.
select a.order_id, a.unit_price, b.order_date
from order_details a full join orders b on a.order_id = b.order_id;
select a.order_id, b.product_id
from orders a full join products b on a.order_id = b.product_id
where product_id is null;

-- combine 3 tables 
select a.first_name, a.last_name, b.order_date, avg(c.amount)
from customers a inner join orders b on a.customer_id = b.customer_id
left outer join payments c on b.order_id = c.order_id
group by a.first_name, a.last_name, b.order_date, c.amount
order by a.first_name;

-- subqueries ()
select * from (
      select * from payments 
      where payment_date >= '2026-07-01'
      order by payment_id
) a
where a.amount = 42000; 

--                                   FUNCTION
                                         
-- order by and group by vs over()
-- over its show the total amount of the table or total sales
SELECT
    order_id,
    amount,
    SUM(amount) AS total_amount
FROM payments
group by order_id, amount
order by order_id, amount;

SELECT
    order_id,
    amount,
    SUM(amount) OVER () AS total_amount
FROM payments;

-- To see how many days before the order was canceled.
select order_id, (order_date - cancel_date) as los,
avg(order_date - cancel_date) over() avg_los
from orders
order by order_id;
select * from orders;

-- avg and round function
with orders_los as 
 (select order_id, (order_date - cancel_date) as los,
  avg(order_date - cancel_date) over() avg_los
  from orders
order by order_id
)
select *,
  round(los - avg_los, 2) as round_los
from orders_los;
--------------------------- Always based on partition by -------------------
-- the rank restarts at 1 for every different amount.
-- RANK() OVER (PARTITION BY) → divides rows for a calculation.
-- (partiton by amount) means you can see different amount.
select 
  order_id, payment_date, amount,
  rank() over(partition by amount order by payment_date desc ) as payment_rank
  from payments;

-- i use WINDOW w AS (PARTITION BY a.order_id ) to do it once
-- instead of SUM(price) OVER (PARTITION BY a.order_id) and SUM(amount) OVER (PARTITION BY a.order_id)
select  
 a.payment_id,
 a.amount,
 avg (amount) over w as avg_amount,
 b.price,
 sum (price) over w as sum_price
from payments a
left outer join products b on a.payment_id = b.product_id
window w as (partition by a.order_id);

-- ROW_NUMBER()  → 1, 2, 3, 4
-- RANK()        → 1, 1, 3, 4 
-- DENSE_RANK()  → 1, 1, 2, 3


select product_id, price,
row_number() over (order by price) as row_price
from products;
select product_id, price,
rank() over (order by price) as rank_price
from products;
select product_id, price,
dense_rank() over (order by price) as dense_price
from products;
-- 
select  
 a.payment_id,
 a.amount,
 row_number() over (partition by a.amount order by a.payment_id asc) as row_cost,
 b.product_name, b.price,
 rank() over (partition by b.price order by b.product_name desc) as rank_cost
from payments a
left outer join products b on a.payment_id = b.product_id;

-- means row_num is base on payment_method
select payment_method, payment_id, amount,
row_number() over(partition by payment_method
order by amount) as row_num
from payments;

-- window function
-- row_number() over(partition by amount order by payment_method)
-- means row_num is base on amount
select payment_method, payment_id, amount,
row_number() over(partition by amount
order by payment_method) as row_num
from payments;

-- rank 
select payment_method, payment_id, amount,
rank () over(partition by payment_method
order by amount) as row_num
from payments;

-- dense_rank
select payment_method, payment_id, amount,
dense_rank () over(partition by payment_method
order by amount) as row_num
from payments;

-- without partition 
-- Without partitioning, the result doesn't make sense.
select payment_method, payment_id, amount,
row_number() over() as row_num
from payments;


-- lag look forward
-- lag look backward
select product_id, price,
lag (price) over (order by product_id asc) as lag_price,
 product_id,
lead (price) over (order by product_id desc) as lead_price
from products;

select product_id, stock, price,
lag (price) over (partition by stock order by product_id asc) as lag_price,
lead (price) over (partition by stock order by product_id desc) as lead_price
from products;


-- FOREIGN KEY must reference a primary key or unique key.
alter table orders
add constraint orders_customers_fk
foreign key (customer_id)
references customers(customer_id);

-- Parent Payments, Child customer_id.

alter table customers
add constraint payment_customers_fk
foreign key (customer_id)
references payments(payment_id);
F
select * from information_schema.table_constraints
where table_schema = 'public';

-- amount range
select max(amount) - min(amount) as amount_range
from payments;

-- Greater than 2,250, and the date is 2026-07-01.
select * from payments
where amount > 2250 and payment_date = '2026-07-01';

-- using in
select * from payments
where payment_id in (1, 2, 3

-- using Like use this %A% if you want letter A annywhere.
select * from suppliers
where supplier_name like 'a%';

-- asc lowest to highest
select contact_name, supplier_id 
from suppliers
order by supplier_id desc;
select contact_name, supplier_id 
from suppliers
order by supplier_id asc;
-- limit 
select contact_name, supplier_id
from suppliers
limit 10;

-- case, when, then, else function
select order_id, payment_date,
case when payment_date = '2026-07-01' then 'credit card' 
     when payment_date = '2026-07-02' then 'GCash'
     else 'cash' end as payment_method
from payments;

-- max,avg function
select avg(amount) as avg_amount, max(amount) as max_amount
from payments;

-- subtract max - min as amount_range
select min(amount) - max(amount) as amount_range
from payments;

-- avg, having function
select payment_id, payment_method, avg(amount) as avg_amount
from payments
where payment_method = 'Cash'
group by payment_id
having avg(amount) < 1250
order by payment_id;

-- min, having fuction
select product_id, stock, product_name, min(price) as min_price
from products
where stock >= 6
group by product_id
having min(price) < 12000
order by product_id;

-- count, distinct function
-- 5 payment_method
select count(distinct payment_method)
from payments;
-- and 
select * from products;
where payment_date '2026-07-01' and payment_method = 'cash';
-- null
select * from payments
where payment_method is null;

                  ----- Subqueries in the WHERE & HAVING Clauses

-- subqeries()
select product_name,
       price - (select avg(price) from products) as diff_from_price
from products;

-- two culumn with one table.
-- subquery that calculates the average price for each product_id.
select * 
     from products a left join(
     select product_id, avg(price) as avg_price
     from products 
     group by product_id) as product_b
on a.product_id = product_b.product_id
where stock >= 60;
-- max 
select * from payments a left join(
         select payment_id, max(amount) as max_mount
		 from payments
		 group by payment_id) as payment_b
		 on a.payment_id = payment_b.payment_id;
-- max and ledt join		 
select
       a.order_id, a.order_date, a.cancel_date, b.max_amount
from 
(select order_id, order_date, cancel_date from orders) as a
left join
         (select order_id,
	     max(amount) as max_amount
		 from payments b
		 group by order_id) as b
		 on a.order_id = b.order_id;
-- above average and where
select * from products
where price > (select avg(price) from products);
-- avg with group by
select product_name, avg(price) as avg_price
from products
group by product_name;

-- above average for each product_name (having)
select product_name, avg(price) as avg_price
from products
group by product_name
having avg(price) > (select avg(price) from products);

-- (any) price that are greater than any of amount from payments
select * from products
where price >
any (select amount from payments);
-- (all)
select * from products
where price >
all (select amount from payments
     where amount = 45000);

select * from products;
select * from payments;

-- EXISTS checks whether the same column value exists and you can also use inner join to see is the same.
select * from orders b
where exists (select a.order_id from payments a
              where a.order_id = b.order_id);
-- With CTE to see above average
with avg_products as 
     (select product_name, avg(price) as avg_price
      from products
	  group by product_name)
select product_name, avg_price
       from avg_products
	   where avg_price > 2800;
	   
-- total Value amount * price	   
select b.product_name, b.price, a.amount, a.payment_method,
       a.amount * b.price as total_value
from payments a
inner join products b on a.payment_id = b.product_id;

-- i use having to see all total_value over 810000
select b.product_id,
       sum(a.amount * b.price) as total_value
from payments a
inner join products b on a.payment_id = b.product_id
group by b.product_id
having SUM(a.amount * b.price) >= 810000
order by total_value;

-- with cte to count
with TV as (select b.product_id,
       sum(a.amount * b.price) as total_value
       from payments a
       inner join products b on a.payment_id = b.product_id
       group by b.product_id
       having sum(a.amount * b.price) >= 810000
       order by total_value)
select * from tv
by order by total_value;

-- aggregate function
select payment_method,
avg(amount) as avg_amount
from payments
group by payment_method;

-- First_Value
select 
  payment_method,amount,
  first_value(amount) over(partition by amount 
  order by payment_method desc) as firts_payment
  from payments
-- last_value
select 
  payment_method,amount,
  last_value(amount) over(partition by amount 
  order by payment_method desc) as firts_payment
  from payments;
  
-- nth_value
-- number 2 means second value or duplice if 3 trplet
select 
  payment_method,amount,
  nth_value(payment_method, 2) over(partition by amount 
  order by payment_method desc) as firts_payment
  from payments;

-- First_Value with cte
with FP_Value as (select 
  payment_method,amount,
  first_value(amount) over(partition by amount 
  order by payment_method desc) as firts_payment
  from payments) 
select * 
from FP_Value
where amount = 500;

-- NTILE
select product_id, product_name, price, stock,
ntile(4) over(partition by stock order by price desc) as ntile_payments
from products
where stock = 10;

-- Amount of payments of each order 
select a.order_id, a.order_date, a.cancel_date, sum(b.amount)
from orders a left join payments b
on a.order_id = b.order_id
group by a.order_id;

-- date
select payment_id, payment_date,
  extract(year from payment_date) as_year,
  extract(month from payment_date) as_date
from payments;

-- i Create table of My_events for date function
create table my_events (
   event_name varchar(50),
   event_date date,
   event_datetime timestamp,
   event_type varchar(20),
   event_desc text);

select event_name, event_date, event_datetime,
      extract (year from event_date) as e_year,
	  extract (month from event_date) as e_month,
	  extract (dow from event_date) as e_dow
	  from my_events;

-- 	extract date and put 1 to 7 case
with date_e as (select event_name, event_date, event_datetime,
      extract (year from event_date) as e_year,
	  extract (month from event_date) as e_month,
	  extract (dow from event_date) as e_day
	  from my_events)

select *, case when e_day = 1 then 'sunday'
               when e_day = 2 then 'monday'
		       when e_day = 3 then 'teusday'
		       when e_day = 4 then 'wednesday'
		       when e_day = 5 then 'thursday'
		       when e_day = 6 then 'friday'
		       when e_day = 7 then 'suturday'
		  else 'unknown' end as event_dow_name
from date_e;

-- difference between event_date and current_date
select event_name, event_date, event_datetime, current_date,
       event_date - current_date as days_until
from my_events;

-- add one hour to event_datetime, use intervel + 1
select event_name, event_date, event_datetime, 
       event_datetime + interval '1 hour' as plus_one_hour
from my_events;	

-- extract 2026 january to april
select * from my_events
where extract(year from event_date) = 2026 and 
extract(month from event_date) between 1 and 4;

-- subtract to days fro preparation_date
select event_name, event_date, event_datetime, 
       event_date - interval '2 day' as Preparation_date
from my_events;	

-- Big letter to small letter
select event_name, upper(event_name), lower(event_name)
from my_events;

-- REPLACE(TRIM): first '' is what to erase, second '' is what to replace it with
-- length count the string including the space
select event_name, event_type,
       replace(trim(event_type), 'p', '') as et_clean,
	   event_desc, length(event_desc) as desc_len
from my_events;

-- length
select length(event_desc)
from my_events;

-- cte with length and concat
with my_events_clean as (select event_name, event_type,
       replace(trim(event_type), 'p', '') as et_clean,
	   event_desc, length(event_desc) as desc_len
from my_events)

select event_name, et_clean, event_desc,
concat(et_clean, ' ! ', event_desc) as event_details
from my_events_clean;

-- concat
select event_name, event_desc,
concat(event_name, ' wow ', event_desc) AS event_details
from my_events;

-- replace
select supplier_id, supplier_name,
replace(supplier_name, 'ABC Electronics', 'abc') as inc_remove
from suppliers
group by supplier_id, supplier_name
order by supplier_id;

-- replace(trim) only remove
select supplier_id, supplier_name,
replace(trim(supplier_name), 'ABC', '') as replace_trim
from suppliers
group by supplier_id, supplier_name
order by supplier_id;

-- with cte
with ABC as (select supplier_id, supplier_name,
replace(trim(supplier_name), 'ABC', '') as replace_trim
from suppliers
group by supplier_id, supplier_name
order by supplier_id)

select supplier_id, supplier_name,
concat(supplier_id, 'for', supplier_name) as supplier_IDN
from ABC;
                         ------------- String-----------------
-- substr 1, 3 means take first 3 letter in event_name 
select event_name,
substr(event_name, 1, 3)
from my_events;

-- counts the characters until the first space
SELECT 
    event_name,
    STRPOS(event_name, ' ')
FROM my_events;

-- get the first word
select event_name,
substr(event_name, 1, STRPOS(event_name,' ')) as firts_word
from my_events;

-- get all the first_word use zero if one the first string is space
select event_name,
case when STRPOS(event_name, ' ') = 0 then event_name
     else substr(event_name, 1, STRPOS(event_name,' ') - 1 ) end as first_word
from my_events;

-- strpos() and replace
select supplier_id, supplier_name,
substr(supplier_name, STRPOS(supplier_name, ' ') + 1),
replace(supplier_name, 'ABC Electronics', 'abc') as inc_remove
from suppliers
group by supplier_id, supplier_name
order by supplier_id;

---------------------------------------- null function -------------------------------------

select * from a
where cancel_date is not null;

-- is null
select order_date, cancel_date 
from orders
where cancel_date is null;

-- is not null
select order_date,cancel_date
from orders
where cancel_date is not null;

-- case and null 
-- need :: text to take effect the else 'return'
-- ::TEXT converts a value into text (string) so it can be treated as text.
select cancel_date, order_date,
case when cancel_date is not null then cancel_date::TEXT
     else 'return' end as return_products
from orders;	

-- coalesce, It is commonly used to replace NULL values
select cancel_date, order_date,
coalesce(cancel_date::TEXT, 'return') as return_p
from orders;

-- null with cte
select cancel_date, order_date, count(order_status) as num_status
from orders
where cancel_date is null
group by cancel_date, order_date;

select cancel_date, order_date, count(order_status) as num_status
             from orders
             where cancel_date is null
             group by cancel_date, order_date;

with num as (select cancel_date, order_date, count(order_status) as num_status
             from orders
             where cancel_date is null
             group by cancel_date, order_date)
			 
             select cancel_date, order_date,
             row_number() over(partition by cancel_date order by num_status desc) as nums
             from num;

with num as (select cancel_date, order_date, count(order_status) as num_status
             from orders
             where cancel_date is null
             group by cancel_date, order_date),
			 
    nums as  (select cancel_date, order_date,
             row_number() over(partition by cancel_date order by num_status desc) as nums
             from num)
select cancel_date, order_date
from nums
where nums = 1;

------------------------------- Dublicate---------------------------------------------
-- view dublicate
select order_status, count(*) as num_status
from orders
group by order_status;

-- using having 
select order_status, count(*) as num_status
from orders
group by order_status
having count(*) > 1;

-- having and order_date to see the fully dublicate
select order_date, order_status, count(*) as num_status
from orders
group by order_date, order_status
having count(*) > 1;

select cancel_date, order_status, count(*) as num_cancel_date
from orders
group by cancel_date, order_status
having count(*) > 1;

-- with order_date to see if have a same date and order_status
select order_date, order_status, count(*) as num_status
from orders
group by order_date, order_status;

-- to see how many dublicate
select distinct payment_method, amount,
row_number() over(partition by payment_method order by amount desc) as row_num
from payments;

-- put subquries to see the specific dublicate
select * from (select distinct payment_method, amount,
row_number() over(partition by payment_method order by amount desc) as row_num
from payments) as num
where row_num = 12;

-- row the number of duplicate payment_method
select *, 
row_number() over(partition by payment_method order by payment_id desc) as count_payment_m
from payments;

-- with cte to see the specific dublicate
with CP as (select *, 
row_number() over(partition by payment_method order by payment_id desc) as count_payment_m
from payments)
select * from cp
where count_payment_m = 4;

------------------------------------ min and max ---------------------------

-- latest payment from different payment_method
select payment_method, max(payment_date) as latest_payment
from payments
group by payment_method; 

-- with amount
select payment_method, max(payment_date) as latest_payment,
       max(amount) max_amount
       from payments
       group by payment_method;

-- left join same table to compare 
with LP as (select payment_method, max(payment_date) as latest_payment_date
            from payments
            group by payment_method)

select * from LP left join payments a
         on lp.payment_method = a.payment_method
		 and lp.latest_payment_date = a.payment_date;

-- to see the same max from amount and price using cte
with PA as (select a.payment_id, b.product_name, max(amount) as max_amount,
            max(price) as max_price
            from payments a left join products b
            on a.payment_id = b.product_id
            group by a.payment_id, b.product_name
            order by a.payment_id) 
			
select max_price, max_amount
from PA
where max_amount = 42000;

-- max price and amount
select a.payment_id, b.product_name, max(amount) as max_amount,
            max(price) as max_price
            from payments a left join products b
            on a.payment_id = b.product_id
            group by a.payment_id, b.product_name
            order by a.payment_id;

select * from 
(select a.payment_method, a.payment_id, b.product_name, max(amount) as max_amount,
            max(price) as max_price,
			row_number() over(partition by payment_method order by product_name) as num_payment
            from payments a left join products b
            on a.payment_id = b.product_id
            group by a.payment_id, b.product_name
            order by a.payment_id, max_price)
where num_payment = 2;
	
-- group by payment_method to see how many sales and to see max price and amount
select a.payment_method, count(a.payment_method), avg(amount) as max_amount,
            max(price) as max_price
            from payments a left join products b
            on a.payment_id = b.product_id
			group by a.payment_method;
			
-- with cte to see stock and the name of product
with PM as (select a.payment_method, count(a.payment_method), max(amount) as max_amount,
            max(price) as max_price
            from payments a left join products b
            on a.payment_id = b.product_id
			group by a.payment_method)
select * from pm
            left join products b 
            on pm.max_amount = b.price
where stock = 22;

-- create pizza table for pivoting
CREATE TABLE pizza_table (
    category varchar(50),
    crush_type VARCHAR(50),
    pizza_name VARCHAR(50),
    price DECIMAL(10, 2)
);

INSERT INTO pizza_table (category, crush_type, pizza_name, price)
VALUES
('Classic', 'Standard Crust', 'Pepperoni Pizza', 299.00),
('Classic', 'Standard Crust', 'Cheese Pizza', 279.00),
('Classic', 'Standard Crust', 'Hawaiian Pizza', 319.00),
('Special', 'Standard Crust', 'Meat Lovers Pizza', 399.00),
('Special', 'Thin Crust', 'BBQ Chicken Pizza', 379.00),
('Vegetarian', 'Thick Crust', 'Margherita Pizza', 259.00),
('Vegetarian', 'Thin Crust', 'Vegetable Pizza', 299.00),
('Premium', 'Stuffed Crust', 'Seafood Pizza', 459.00),
('Premium', 'Thin Crust', 'Supreme Pizza', 429.00),
('Premium', 'Thick Crust', 'Bacon Pizza', 389.00);

select * from pizza_table;

-- then 1 to put one in standard crust else 0 to not Standard Crust
select *,
       case when crush_type = 'Standard Crust' then 1 
	   else 0 end as "Standard Crust"
from pizza_table;

-- summary table
-- to see how many comparison between category and crush_type
select category,
       sum(case when crush_type = 'Standard Crust' then 1 else 0 end) as Standard_Crust,
	   sum(case when crush_type = 'Thin Crust' then 1 else 0 end) as Thin_Crust,
	   sum(case when crush_type = 'Thick Crust' then 1 else 0 end) as Thick_Crust
from pizza_table
group by category;

-- left join and case
select * from payments a left join products b
       on a.payment_id = b.product_id;

-- remove else 0 if you want a null result
-- ROUND() is used to reduce or control the number of decimal places.
select a.payment_method, a.amount, b.price,
       case when stock >= 20 then 1 else 0 end as s_over_20,
	   case when stock >= 22 then 1 else 0 end as s_over_22,
	   case when stock >= 40 then 1 else 0 end as s_over_40,
	   case when stock >= 90 then 1 end as s_over_90,
	   round(avg(case when amount >= 40000 then 1 end)) as over_40000,
	   avg(case when amount >= 1900 then 1 end) as over_1900
       from payments a left join products b
       on a.payment_id = b.product_id
	   group by a.payment_method, a.amount, b.stock, b.price;

------------------------------------------Calculation--------------------------------------------
-- rollup to see the grandtotal
select product_name, stock , sum(price) as total_sales
from products
group by rollup (product_name, stock);

-- with CTE 
with TS as (select product_name, stock , sum(price) as total_sales
            from products
            group by product_name, stock)

select product_name, stock,
       row_number() over(partition by product_name order by stock) 
	   from TS
	   where stock >= 5;
	   
-- Use a CTE to see the total sum for each product_name
with TS as (select product_name, price, stock , sum(price) as total_sales
            from products
            group by product_name, price, stock)

select product_name,price,
       sum(total_sales) 
from ts
group by product_name, price;

-- cte sum and order by to calculate a running total.
-- action movie + acoustic guitar = 17500.00
with TS as (select product_name, price, stock , sum(price) as total_sales
            from products
            group by product_name, price, stock)
			
select product_name, price, stock,		
sum(total_sales) over(order by product_name ) as running_sum
from ts;

























