-- =====================================================================
-- Task 1 - SQL queries (standard SQL: MySQL, PostgreSQL, SQL Server, SQLite)
-- =====================================================================

-- ---------------------------------------------------------------------
-- Q1: Every employee with their department name (LEFT JOIN + table aliases)
--     employees(emp_id, emp_name, dept_id)   departments(dept_id, dept_name)
--
-- employees is the LEFT table, so every employee is kept. If there is no
-- matching department the department columns are NULL. That covers:
--   * Meera : dept_id IS NULL                      -> nothing to match
--   * Karan : dept_id = 30, absent from departments -> no match
-- ---------------------------------------------------------------------
SELECT
    e.emp_name,
    d.dept_name
FROM employees        AS e
LEFT JOIN departments AS d
       ON e.dept_id = d.dept_id;

-- Expected output
-- emp_name | dept_name
-- Asha     | Sales
-- Ravi     | Engineering
-- Meera    | NULL
-- Karan    | NULL


-- ---------------------------------------------------------------------
-- Q2: Total order amount per customer; 0 when a customer has no orders;
--     sorted by total descending
--     orders(order_id, customer_id, order_date, amount)
--     customers(customer_id, customer_name, city)
--
-- customers is the LEFT table so customers without orders survive the join.
-- SUM() of no rows is NULL, so COALESCE converts it to 0.
-- ---------------------------------------------------------------------
SELECT
    c.customer_name,
    COALESCE(SUM(o.amount), 0) AS total_amount
FROM customers   AS c
LEFT JOIN orders AS o
       ON c.customer_id = o.customer_id
GROUP BY c.customer_id, c.customer_name
ORDER BY total_amount DESC;

-- Expected output
-- customer_name | total_amount
-- Asha          | 4300     (2500 + 1800)
-- Ravi          | 4000
-- Karan         | 1200
-- Meera         | 0        (no orders)
