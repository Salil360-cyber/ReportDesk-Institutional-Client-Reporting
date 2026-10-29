-- =====================================================================
-- ReportDesk: SQL solution file (PostgreSQL)
-- Run step by step in pgAdmin. Each step builds on the one before.
-- Run sql/01_setup.sql first.
-- Reporting month: September 2026 (month-end 2026-09-30, previous 2026-08-31)
-- =====================================================================

-- ---------------------------------------------------------------------
-- STEP 1. Value every holding at both month-ends
-- LEFT JOIN keeps holdings even when the security or the price is missing,
-- so the data-quality checks in Step 2 can find them.
-- ---------------------------------------------------------------------
DROP VIEW IF EXISTS v_valuation CASCADE;
CREATE VIEW v_valuation AS
SELECT h.client_id,
       h.security_id,
       s.security_name,
       s.asset_class,
       h.as_of_date,
       h.quantity,
       p.price,
       h.quantity * p.price AS market_value
FROM holdings h
LEFT JOIN securities s ON s.security_id = h.security_id
LEFT JOIN prices     p ON p.security_id = h.security_id
                      AND p.price_date  = h.as_of_date;

-- ---------------------------------------------------------------------
-- STEP 2. Data-quality checks, all in one view
-- BLOCKER = report must not be sent until fixed
-- WARNING = report can go, but the analyst must confirm the number first
-- ---------------------------------------------------------------------
DROP VIEW IF EXISTS v_dq_exceptions CASCADE;
CREATE VIEW v_dq_exceptions AS

-- 2.1 Held security has no month-end price
SELECT client_id, 'Missing month-end price' AS check_name, 'BLOCKER' AS severity,
       security_id, 'No price on ' || as_of_date AS detail
FROM v_valuation
WHERE as_of_date = DATE '2026-09-30' AND asset_class IS NOT NULL AND price IS NULL

UNION ALL
-- 2.2 Holding points to a security that is not in the security master
SELECT client_id, 'Unknown security', 'BLOCKER', security_id, 'Not in securities table'
FROM v_valuation
WHERE as_of_date = DATE '2026-09-30' AND asset_class IS NULL

UNION ALL
-- 2.3 Negative quantity
SELECT client_id, 'Negative quantity', 'BLOCKER', security_id, 'Quantity ' || quantity
FROM holdings
WHERE as_of_date = DATE '2026-09-30' AND quantity < 0

UNION ALL
-- 2.4 Same security booked more than once for a client on the same date
SELECT client_id, 'Duplicate holding row', 'BLOCKER', security_id, COUNT(*) || ' rows'
FROM holdings
WHERE as_of_date = DATE '2026-09-30'
GROUP BY client_id, security_id
HAVING COUNT(*) > 1

UNION ALL
-- 2.5 Cash must reconcile: end cash = start cash + net flows in the month
SELECT c.client_id, 'Cash does not reconcile', 'BLOCKER', 'CASH_USD',
       'Difference ' || (COALESCE(e.qty, 0) - COALESCE(b.qty, 0) - COALESCE(f.net, 0))
FROM clients c
LEFT JOIN (SELECT client_id, SUM(quantity) qty FROM holdings
           WHERE security_id = 'CASH_USD' AND as_of_date = DATE '2026-08-31' GROUP BY client_id) b ON b.client_id = c.client_id
LEFT JOIN (SELECT client_id, SUM(quantity) qty FROM holdings
           WHERE security_id = 'CASH_USD' AND as_of_date = DATE '2026-09-30' GROUP BY client_id) e ON e.client_id = c.client_id
LEFT JOIN (SELECT client_id, SUM(amount) net FROM cash_flows
           WHERE flow_date BETWEEN DATE '2026-09-01' AND DATE '2026-09-30' GROUP BY client_id) f ON f.client_id = c.client_id
WHERE COALESCE(e.qty, 0) - COALESCE(b.qty, 0) - COALESCE(f.net, 0) <> 0

UNION ALL
-- 2.6 Benchmark return missing for the month
SELECT c.client_id, 'Missing benchmark return', 'BLOCKER', c.benchmark_id, 'No return for 2026-09-30'
FROM clients c
LEFT JOIN benchmark_returns br ON br.benchmark_id = c.benchmark_id AND br.month_end = DATE '2026-09-30'
WHERE br.benchmark_id IS NULL

UNION ALL
-- 2.7 Stale price: month-end price exactly equal to last month (cash excluded)
SELECT e.client_id, 'Stale price (unchanged)', 'WARNING', e.security_id, 'Price ' || e.price || ' same as last month'
FROM v_valuation e
JOIN v_valuation b ON b.client_id = e.client_id AND b.security_id = e.security_id
                  AND b.as_of_date = DATE '2026-08-31'
WHERE e.as_of_date = DATE '2026-09-30' AND e.asset_class <> 'Cash' AND e.price = b.price

UNION ALL
-- 2.8 Price moved more than 25% in one month: confirm before reporting
SELECT e.client_id, 'Price move over 25%', 'WARNING', e.security_id,
       'Move ' || ROUND((e.price / b.price - 1) * 100, 1) || '%'
FROM v_valuation e
JOIN v_valuation b ON b.client_id = e.client_id AND b.security_id = e.security_id
                  AND b.as_of_date = DATE '2026-08-31'
WHERE e.as_of_date = DATE '2026-09-30' AND ABS(e.price / b.price - 1) > 0.25;

-- Look at all exceptions
SELECT * FROM v_dq_exceptions ORDER BY severity, client_id;

-- ---------------------------------------------------------------------
-- STEP 3. Is each client's report ready to send?
-- ---------------------------------------------------------------------
DROP VIEW IF EXISTS v_client_readiness CASCADE;
CREATE VIEW v_client_readiness AS
SELECT c.client_id,
       c.client_name,
       COUNT(*) FILTER (WHERE d.severity = 'BLOCKER') AS blockers,
       COUNT(*) FILTER (WHERE d.severity = 'WARNING') AS warnings,
       CASE WHEN COUNT(*) FILTER (WHERE d.severity = 'BLOCKER') > 0 THEN 'BLOCKED'
            WHEN COUNT(*) FILTER (WHERE d.severity = 'WARNING') > 0 THEN 'READY - CHECK WARNINGS'
            ELSE 'READY' END AS report_status
FROM clients c
LEFT JOIN v_dq_exceptions d ON d.client_id = c.client_id
GROUP BY c.client_id, c.client_name;

SELECT * FROM v_client_readiness ORDER BY client_id;

-- ---------------------------------------------------------------------
-- STEP 4. Portfolio summary and performance vs benchmark
-- Return method: simple Dietz = (EMV - BMV - Net Flows) / (BMV + 0.5 x Net Flows)
-- Excess return in basis points (1 bp = 0.01%)
-- ---------------------------------------------------------------------
DROP VIEW IF EXISTS v_portfolio_summary CASCADE;
CREATE VIEW v_portfolio_summary AS
WITH mv AS (
    SELECT client_id,
           SUM(market_value) FILTER (WHERE as_of_date = DATE '2026-08-31') AS bmv,
           SUM(market_value) FILTER (WHERE as_of_date = DATE '2026-09-30') AS emv
    FROM v_valuation
    GROUP BY client_id
),
fl AS (
    SELECT client_id, SUM(amount) AS net_flows
    FROM cash_flows
    WHERE flow_date BETWEEN DATE '2026-09-01' AND DATE '2026-09-30'
    GROUP BY client_id
)
SELECT c.client_id,
       c.client_name,
       c.client_type,
       c.country,
       b.benchmark_name,
       ROUND(mv.bmv, 2)                                   AS bmv,
       ROUND(mv.emv, 2)                                   AS emv,
       COALESCE(fl.net_flows, 0)                          AS net_flows,
       ROUND(mv.emv - mv.bmv - COALESCE(fl.net_flows, 0), 2) AS investment_gain,
       ROUND((mv.emv - mv.bmv - COALESCE(fl.net_flows, 0))
             / (mv.bmv + 0.5 * COALESCE(fl.net_flows, 0)), 6) AS portfolio_return,
       br.return_pct                                      AS benchmark_return,
       ROUND(((mv.emv - mv.bmv - COALESCE(fl.net_flows, 0))
             / (mv.bmv + 0.5 * COALESCE(fl.net_flows, 0)) - br.return_pct) * 10000, 1) AS excess_bps,
       r.report_status
FROM clients c
JOIN mv                      ON mv.client_id = c.client_id
LEFT JOIN fl                 ON fl.client_id = c.client_id
JOIN benchmarks b            ON b.benchmark_id = c.benchmark_id
LEFT JOIN benchmark_returns br ON br.benchmark_id = c.benchmark_id AND br.month_end = DATE '2026-09-30'
JOIN v_client_readiness r    ON r.client_id = c.client_id;

SELECT * FROM v_portfolio_summary ORDER BY client_id;

-- ---------------------------------------------------------------------
-- STEP 5. Asset allocation at month-end (% of portfolio)
-- ---------------------------------------------------------------------
DROP VIEW IF EXISTS v_asset_allocation CASCADE;
CREATE VIEW v_asset_allocation AS
SELECT client_id,
       asset_class,
       ROUND(SUM(market_value), 2) AS market_value,
       ROUND(SUM(market_value) / SUM(SUM(market_value)) OVER (PARTITION BY client_id), 6) AS weight
FROM v_valuation
WHERE as_of_date = DATE '2026-09-30' AND asset_class IS NOT NULL
GROUP BY client_id, asset_class;

SELECT * FROM v_asset_allocation ORDER BY client_id, asset_class;

-- ---------------------------------------------------------------------
-- STEP 6. Top 5 holdings per client (cash excluded)
-- ---------------------------------------------------------------------
DROP VIEW IF EXISTS v_top_holdings CASCADE;
CREATE VIEW v_top_holdings AS
SELECT *
FROM (
    SELECT v.client_id,
           ROW_NUMBER() OVER (PARTITION BY v.client_id ORDER BY v.market_value DESC) AS rank_no,
           v.security_name,
           v.asset_class,
           ROUND(v.market_value, 2) AS market_value,
           ROUND(v.market_value / t.total_mv, 6) AS weight
    FROM v_valuation v
    JOIN (SELECT client_id, SUM(market_value) AS total_mv FROM v_valuation
          WHERE as_of_date = DATE '2026-09-30' GROUP BY client_id) t ON t.client_id = v.client_id
    WHERE v.as_of_date = DATE '2026-09-30' AND v.asset_class <> 'Cash' AND v.market_value IS NOT NULL
) x
WHERE rank_no <= 5;

SELECT * FROM v_top_holdings ORDER BY client_id, rank_no;

-- ---------------------------------------------------------------------
-- STEP 7. Delivery against contractual deadlines (June to August, sent reports)
-- ---------------------------------------------------------------------
SELECT report_month,
       COUNT(*)                                              AS reports_sent,
       COUNT(*) FILTER (WHERE sent_date <= deadline_date)    AS on_time,
       ROUND(100.0 * COUNT(*) FILTER (WHERE sent_date <= deadline_date) / COUNT(*), 1) AS on_time_pct
FROM report_tracker
WHERE status = 'Sent'
GROUP BY report_month
ORDER BY report_month;

-- September reports: deadline status as of business date 2026-10-09
SELECT t.client_id, c.client_name, t.analyst, t.deadline_date, r.report_status,
       CASE WHEN t.deadline_date < DATE '2026-10-09' THEN 'PAST DEADLINE'
            WHEN t.deadline_date = DATE '2026-10-09' THEN 'DUE TODAY'
            ELSE 'UPCOMING' END AS deadline_status
FROM report_tracker t
JOIN clients c ON c.client_id = t.client_id
JOIN v_client_readiness r ON r.client_id = t.client_id
WHERE t.report_month = DATE '2026-09-30'
ORDER BY t.deadline_date, t.client_id;

-- ---------------------------------------------------------------------
-- STEP 8. Export for Excel and Power BI
-- In pgAdmin: run each SELECT, then click "Save results to file" (download icon)
-- and save as CSV with the file name shown.
-- ---------------------------------------------------------------------
-- portfolio_summary.csv
SELECT * FROM v_portfolio_summary ORDER BY client_id;
-- asset_allocation.csv
SELECT * FROM v_asset_allocation ORDER BY client_id, asset_class;
-- top_holdings.csv
SELECT * FROM v_top_holdings ORDER BY client_id, rank_no;
-- dq_exceptions.csv
SELECT * FROM v_dq_exceptions ORDER BY severity, client_id;
-- client_readiness.csv
SELECT * FROM v_client_readiness ORDER BY client_id;
-- report_tracker.csv
SELECT * FROM report_tracker ORDER BY report_month, client_id;
-- client_queries.csv
SELECT * FROM client_queries ORDER BY received_date;
