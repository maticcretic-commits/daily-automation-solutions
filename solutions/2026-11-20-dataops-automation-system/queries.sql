-- Auditable SQL for the DataOps pipeline (pipeline.py executes these
-- exact statements against the SQLite database built from data/*.csv).
-- month is derived with substr(date, 1, 7) because dates are ISO text.

-- QUERY: monthly
SELECT substr(date, 1, 7) AS month,
       ROUND(SUM(qty * unit_price), 2) AS revenue,
       ROUND(SUM(qty * unit_cost), 2) AS cogs,
       ROUND(SUM(qty * (unit_price - unit_cost)), 2) AS gross_profit,
       COUNT(*) AS orders
FROM sales
GROUP BY month
ORDER BY month;

-- QUERY: by_category
SELECT category,
       ROUND(SUM(qty * unit_price), 2) AS revenue,
       ROUND(SUM(qty * (unit_price - unit_cost)), 2) AS gross_profit,
       COUNT(*) AS orders
FROM sales
GROUP BY category
ORDER BY revenue DESC;

-- QUERY: by_channel
SELECT channel,
       ROUND(SUM(qty * unit_price), 2) AS revenue,
       COUNT(*) AS orders
FROM sales
GROUP BY channel
ORDER BY revenue DESC;

-- QUERY: expenses_monthly
SELECT substr(date, 1, 7) AS month,
       ROUND(SUM(amount), 2) AS expenses
FROM expenses
GROUP BY month
ORDER BY month;

-- QUERY: purchases_monthly
SELECT substr(date, 1, 7) AS month,
       ROUND(SUM(qty * unit_cost), 2) AS purchases
FROM purchases
GROUP BY month
ORDER BY month;
