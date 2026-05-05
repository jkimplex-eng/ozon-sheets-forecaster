# Next Codex task: Ads debug and Raw_Ads_Daily statistics

## Current status

The Apps Script project is now in the correct repository: `jkimplex-eng/ozon-sheets-forecaster`.

Confirmed working in Google Sheets:

- Products load from Ozon Seller API.
- Stocks load from Ozon Seller API.
- Sales/Finance load from Ozon Seller API.
- `Ad_Campaigns` loads from Ozon Performance API using `GET /api/client/campaign`.
- Summary builds a basic P&L using `Unit_Costs`.

Current problem:

- `Ad_Campaigns` is only a campaign catalog.
- `Raw_Ads_Daily` still does not contain real advertising statistics.
- Summary/P&L still shows advertising spend as `нет данных`.

## Required task

Implement end-to-end advertising statistics diagnostics and loading.

## Requirements

### 1. Add `Ads_Debug` sheet

Headers:

```text
run_at, step, campaign_id, campaign_name, endpoint, http_code, status, message, preview
```

### 2. Add menu item

Add to `onOpen()`:

```text
Ozon -> Debug рекламы
```

It should call:

```javascript
function debugAdsStats()
```

### 3. Debug flow

`debugAdsStats()` should:

1. Read first 3-5 `CAMPAIGN_STATE_RUNNING` campaign IDs from `Ad_Campaigns`.
2. Request statistics for the last 7 days.
3. Write every step to `Ads_Debug`:
   - campaigns selected
   - request endpoint
   - request payload preview
   - report UUID if received
   - polling attempts
   - raw response preview
   - raw stats rows count
   - normalized rows count
   - written rows count
4. If normalized rows > 0, write to `Raw_Ads_Daily`.
5. If normalized rows = 0, do not clear old `Raw_Ads_Daily`.

### 4. Performance statistics flow

Implement or fix:

```javascript
updateAdsStatsDaily(startDate, endDate)
requestPerformanceStatsReport_(campaignIds, startDate, endDate, groupBy)
downloadPerformanceStatsReport_(uuid)
```

Use host:

```text
https://api-performance.ozon.ru
```

Do not use old host:

```text
https://performance.ozon.ru
```

Candidate endpoints for report request:

```text
POST /api/client/statistics/json
```

Candidate endpoints for download/polling:

```text
GET /api/client/statistics/report?UUID=<uuid>
GET /api/client/statistics/report?uuid=<uuid>
GET /api/client/statistics/json?UUID=<uuid>
GET /api/client/statistics/json?uuid=<uuid>
```

Support both cases:

- report request returns UUID and requires polling;
- report request returns rows directly.

### 5. Rate limits

Performance API can return:

```text
Превышен лимит активных запросов (максимум 1)
```

Use sequential requests only:

- batch size: 5-10 campaign IDs
- `Utilities.sleep(1200)` between calls
- retry/backoff for HTTP 429

### 6. Normalize rows into `Raw_Ads_Daily`

Required columns:

```text
date, campaign_id, campaign_name, campaign_type, placement, sku, offer_id, impressions, clicks, ctr, spend, avg_cpc, orders, revenue, drr_acos, roas, search_phrase, source_endpoint, loaded_at
```

Support possible source field names:

- `impressions`, `shows`, `views`, `showCount`
- `clicks`, `clickCount`
- `spend`, `expense`, `cost`, `moneySpent`, `money_spent`
- `orders`, `ordersCount`, `orders_count`
- `revenue`, `sales`, `money`, `attributedRevenue`
- `sku`, `offerId`, `offer_id`

Do not use campaign `budget`, `dailyBudget`, or `weeklyBudget` as spend. They are budget settings, not actual spend.

### 7. Summary diagnostics

In `buildSummary()` add Diagnostics entries:

- `Raw_Ads_Daily headers detected`
- `Raw_Ads_Daily data rows count`
- detected spend column name
- detected impressions column name
- detected clicks column name
- ad spend sum
- impressions sum
- clicks sum
- ad revenue sum
- ad orders sum

If `Raw_Ads_Daily` is empty, write warning:

```text
Raw_Ads_Daily пустая: Summary не может посчитать рекламные расходы
```

If rows exist but spend is zero:

```text
Raw_Ads_Daily есть, но spend не найден или равен 0
```

### 8. Fix low-stock display

Summary currently shows product IDs in low stock rows. It should display `offer_id` when possible.

Build mapping from `Products`:

- `product_id -> offer_id`
- `sku -> offer_id`

In low stock summary rows, use:

1. `offer_id`
2. mapped `offer_id`
3. `sku`
4. `product_id`

Deduplicate by displayed SKU.

### 9. P&L calculation status

Add Summary row:

```text
P&L | Статус расчёта | Неполный | нет данных: commission, logistics, returns, ads
```

If ads spend exists, remove `ads` from missing list.

### 10. README update

Document:

- `Ad_Campaigns` is a campaign catalog.
- `Raw_Ads_Daily` is actual ad statistics.
- If Summary does not show ad spend, check `Ads_Debug`, `Raw_Ads_Daily`, and `Diagnostics`.
- 429 requires retry/backoff.

## Checks

```bash
rg "Ads_Debug|Debug рекламы|Raw_Ads_Daily headers detected|Raw_Ads_Daily rows written|Performance statistics returned 0" Code.gs README.md
rg "buildSummary|updateAdsStatsDaily|parseNumber_|spend|impressions|clicks" Code.gs
rg "product_id.*offer_id|displaySku|Низкие остатки" Code.gs
rg "api-performance.ozon.ru|performance.ozon.ru" Code.gs README.md
```

## Commit message

```text
Add ads debug flow and Summary ad diagnostics
```
