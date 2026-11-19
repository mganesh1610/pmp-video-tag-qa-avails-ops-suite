# Schema & Methodology Specification

## 1. Relational Star Schema (`pmp_avails_tag_ops.db`)

- **`dim_vertical`** (`6` rows): `vertical_id` (`PK`), `vertical_name`, `floor_cpm_usd`.
- **`fact_publisher_articles`** (`39,644` rows): `article_id` (`PK`), `vertical_id` (`FK -> dim_vertical`), `format_bucket` (`Post-Content High-Impact` | `Outstream Video VAST 4.2`), `device_viewport` (`Desktop` | `Mobile`), `is_weekend`, `num_imgs`, `num_videos`, `shares`, `est_monthly_impressions`.
- **`fact_pmp_auctions`** (`200,000` rows): `auction_id` (`PK`), `split_name` (`TRAIN` | `TEST`), `vertical_id` (`FK -> dim_vertical`), `dsp_seat`, `deal_type` (`PMP Deal ID` | `Open Exchange`), `format_bucket`, `device_viewport`, `clicked`, `winning_price_raw`, `bid_floor_cpm`, `adserver_gross_imps`, `dsp_billable_imps`, `billable_spend_micros`, `troubleshoot_flag`.

## 2. Formulations

1. **Contended Bookable PMP Avails:**
   $$\text{Avails}_{\text{bookable}}(v, f, d) = \operatorname{round}\!\left((1 - c) \sum_{i \in S(v,f,d)} I_i\right), \quad c = 0.18$$
   where $c = 0.18$ reserves an 18% contention buffer for overlapping guaranteed and private marketplace line items.

2. **Ad-Server vs. DSP Discrepancy & Exact Micros Spend Conservation:**
   $$\Delta_{\text{imps}} = \sum I_{\text{GAM, gross}} - \sum I_{\text{DSP, billable}}, \qquad \text{Drift}_{\text{USD}} = \frac{\sum M_{\text{SQL}} - \sum_{s \in \text{DSP}} M_s}{10^6} = \$0.00$$
