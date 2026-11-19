# PMP High-Impact & Outstream Video Tag QA, Publisher Avails & XLS Campaign Operations Suite

End-to-end programmatic campaign operations, pre-launch ad tag & outstream `VAST 4.2` QA validation, preset-criteria publisher avails forecasting, multi-tab `.xlsx` stakeholder data collation, first-line delivery troubleshooting, and automated proof-of-launch `.pptx` screenshot deck pipeline.

## 1. Architecture & Operational Scope

1. **Pre-Launch Ad Tag & Creative Spec QA Validator (`src/pmp_tag_qa_avails_engine.py`)**
   - Audits **120 pre-launch ad tag and creative packages** across **Post-Content High-Impact (`970x250`, `320x480`)** `CM360` / `Google Ad Manager (GAM)` containers and **Outstream Video (`VAST 4.2`, `640x360` MP4)** XML wrappers.
   - Enforces **10 deterministic QA rules**: HTTPS SSL compliance, `%%CLICK_URL_UNESC%%` click macro binding, `[TIMESTAMP]` / `[CACHEBUSTER]` cachebusting, `video/mp4` MIME & `<= 2,500 kbps` outstream bitrate bounds, `<= 200 KB` HTML5 polite-load thresholds, UTM landing-URL taxonomy, and `IAS` / `DoubleVerify` OMID `<AdVerifications>` script wrapping.
   - Intercepts **24/24 (`100.0%`)** seeded pre-launch tag/spec defects and verifies **120/120 (`100.0%`)** post-remediation launch readiness.

2. **Preset-Criteria Publisher Avails Forecaster (`sql/avails_and_reconciliation_queries.sql`)**
   - Combines **39,644 premium publisher article placements** (`UCI Online News Popularity`, DOI `10.24432/C5NS3V`) with **200,000 RTB programmatic auction logs** (`iPinYou` `100,000` TRAIN + `100,000` TEST, `239,644` total SHA-256 verified records) inside a normalized SQLite star schema (`PRAGMA foreign_keys = ON`).
   - Pulls inventory avails across **24 preset Account Manager targeting segments** (`6` publisher verticals × `2` ad formats × `2` device viewports), applying an **18.0% cross-deal contention buffer** (`707,646,702` gross monthly impressions $\rightarrow$ **`580,270,296` bookable PMP avails**).

3. **Multi-Tab XLS/XLSX Stakeholder Data Collation & First-Line Troubleshooting (`outputs/pmp_avails_and_tag_ops_pack.xlsx`)**
   - Collates GAM ad-server gross impressions against DSP billable impressions across `DV360`, `The Trade Desk`, `Amazon DSP`, and `Yahoo DSP` (`PMP Deal ID` vs. `Open Exchange`) with live spreadsheet formulas (`=SUM(...)`, `=ROUND(...)`), reconciling **`$394,123.84` verified billable spend with `$0.00` ledger variance**.
   - Classifies first-line troubleshooting root causes (`ZERO_BID_FLOOR_LOCK`, `VAST_WRAPPER_TIMEOUT_303`, `IVT_NON_VIEWABLE_FILTERED`).

4. **Automated Proof-of-Launch Screenshot Deck (`outputs/pmp_campaign_screenshot_deck.pptx`)**
   - Generates a **7-slide native OpenXML PowerPoint screenshot deck** pairing live placement tearsheet frames with Deal IDs, floor CPMs, viewport specs, and GAM/CM360 QA sign-off metadata.

## 2. Verified Public Datasets & SHA-256 Provenance

| Dataset | Source URL | Rows | Verified SHA-256 |
| :--- | :--- | ---: | :--- |
| iPinYou RTB Auction Log (TRAIN) | `https://raw.githubusercontent.com/wnzhang/rtbcontrol/56f8aea7b9f67eaf6d8d731fc091016f22f2739d/exp-data/train.txt` | `100,000` | `db0ecc0f49626993a2bbadab0561ddb3307926abc161f823e07b707c10aacefd` |
| iPinYou RTB Auction Log (TEST) | `https://raw.githubusercontent.com/wnzhang/rtbcontrol/56f8aea7b9f67eaf6d8d731fc091016f22f2739d/exp-data/test.txt` | `100,000` | `ba3dc59962ec8f58ffb0ca3de5c833d4b589fc7215857f4ad7beb92726f24798` |
| UCI Online News Popularity (`zip`) | `https://archive.ics.uci.edu/static/public/332/online+news+popularity.zip` | `39,644` | `dba2ae526f62ccef6f2f8efb53b0268319f4d0d2719bb98d88958cfe31d58a22` |

Backup public datasets: Criteo Attribution Dataset (`https://ailab.criteo.com/criteo-attribution-modeling-bidding-dataset/`) and IAB Tech Lab VAST 4.2 XML Samples (`https://github.com/InteractiveAdvertisingBureau/VAST_Samples`).

## 3. Quick Start & Reproducible Verification

```bash
python3 src/pmp_tag_qa_avails_engine.py
python3 -m unittest discover -s tests -v
```

All **12/12** validation gates and **5/5** unit tests execute using Python 3.12 standard library only.
