-- Preset-Criteria Publisher Avails Pull & Cross-DSP PMP Delivery Reconciliation (SQL Joins & CTEs)

WITH article_supply AS (
    SELECT
        a.vertical_id,
        v.vertical_name,
        v.floor_cpm_usd,
        a.format_bucket,
        a.device_viewport,
        COUNT(*) AS article_placements,
        SUM(a.est_monthly_impressions) AS gross_avails_imps
    FROM fact_publisher_articles a
    JOIN dim_vertical v ON a.vertical_id = v.vertical_id
    GROUP BY a.vertical_id, v.vertical_name, v.floor_cpm_usd, a.format_bucket, a.device_viewport
),
auction_fill AS (
    SELECT
        vertical_id,
        format_bucket,
        device_viewport,
        COUNT(*) AS auction_requests,
        SUM(CASE WHEN troubleshoot_flag = 'HEALTHY_DELIVERY' THEN 1 ELSE 0 END) AS healthy_wins,
        SUM(adserver_gross_imps) AS adserver_imps,
        SUM(dsp_billable_imps) AS billable_imps,
        SUM(billable_spend_micros) AS spend_micros
    FROM fact_pmp_auctions
    WHERE deal_type = 'PMP Deal ID'
    GROUP BY vertical_id, format_bucket, device_viewport
)
SELECT
    s.vertical_id,
    s.vertical_name,
    s.format_bucket,
    s.device_viewport,
    s.floor_cpm_usd,
    s.article_placements,
    s.gross_avails_imps,
    CAST(ROUND(s.gross_avails_imps * (1.0 - 0.18)) AS INTEGER) AS bookable_pmp_avails_imps,
    COALESCE(f.auction_requests, 0) AS pmp_auction_requests,
    ROUND(100.0 * COALESCE(f.healthy_wins, 0) / NULLIF(f.auction_requests, 0), 2) AS pmp_clear_rate_pct,
    ROUND(COALESCE(f.spend_micros, 0) / 1000000.0, 2) AS verified_pmp_spend_usd
FROM article_supply s
LEFT JOIN auction_fill f
  ON s.vertical_id = f.vertical_id
 AND s.format_bucket = f.format_bucket
 AND s.device_viewport = f.device_viewport
ORDER BY s.vertical_id, s.format_bucket, s.device_viewport;
