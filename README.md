# Zalando DE Menswear Pricing & Brand Analytics

Excel analysis of a 999-product scrape of Zalando.de's men's basic t-shirt and polo catalogue (January 2023). Answers 20 structured business/analytics questions covering pricing, discounting, brand performance, product lookup, and dashboarding.

## Dataset

| | |
|---|---|
| Products | 999 |
| Unique brands | 145 |
| Columns | 21 |
| Category | Men's T-Shirts & Polos |
| Price range | €6.99 – €99.99 |
| Scan date | 2023-01-29 |

Raw fields: `SKU, Product_url, Brand, Product_name, Category, Subcategory, Subsubcategory, Color, Size_range, Price_retail, Price_current, Price_label, Currency, Sort_by_ranking, Rating, Reviews_count, Promotion, Sku_variant, In_stock, Scan_date, Image_url`

Two derived fields were added during analysis: `Discount` (calculated markdown %) and `Discount Level` (tiered classification).

## Workbook structure

| Sheet | Contents |
|---|---|
| `Row` | Full product table + calculated `Discount` and `Discount Level` columns |
| `Real Sale` | Full product table with additional helper columns (`Basic Product Name`, `Price range`) used for lookups and text parsing |
| `Pricing & Discount Analysis` | Discount %, discounted product count, per-brand price stats, top discounted brands |
| `Brand` | Per-brand product counts, market share, and revenue potential |
| `Product Lookup & Search` | XLOOKUP-based SKU search tool |
| `Pivot Tables & Dashboards` | Subcategory/sub-subcategory breakdown, price distribution, and discount-by-brand pivots |

## Key findings

- **406 of 999 products (~41%)** are currently discounted.
- **Nike, Adidas Performance, and Tommy Hilfiger** are the top 3 brands by SKU count (100, 75, and 62 products respectively), together representing roughly a quarter of the catalogue.
- Revenue potential is concentrated at the top: Nike alone accounts for **€1,750** in summed current pricing, more than any other brand.
- The steepest average discounts belong to smaller brands — **Produkt (~70%)**, **Zign (~70%)**, and **Selected Homme (~63%)** — while large brands like Nike and Tommy Hilfiger discount less aggressively on average.
- Price distribution is bottom-heavy: **438 products (~44%) sit in the €0–15 band**, versus only 103 products above €50.
- Products break down almost evenly between **T-Shirts (486)** and **Polo Shirts (498)**, with Basic T-Shirts and Pique Polos as the largest sub-subcategories.
- Discounts were classified into four tiers — **No Discount, Light, Moderate, High** — to support the conditional formatting and brand-level discount counts.

## Questions answered

1. Discount percentage per product & biggest markdown
2. Count of discounted products
3. Average/min/max current price per brand (pivot table)
4. Discount tier classification (No Discount / Light / Moderate / Deep)
5. Top 10 biggest price drops
6. Products per brand & brand with most SKUs
7. Brand market share
8. Premium vs. budget brand positioning
9. Top 10 brands bar chart
10. Revenue potential per brand
11. SKU lookup tool (XLOOKUP)
12. Products with "basic" in the name
13. Colour extraction from product name
14. Unique sorted brand list
15. Rank #1 product and brand
16. Subcategory / sub-subcategory breakdown (pivot table)
17. Price distribution histogram
18. Conditional formatting for discounts > 40%
19. Heavily discounted products per brand (>30%)
20. Scan date parsing & day of week

## Skills used

`XLOOKUP` · `INDEX-MATCH` · `COUNTIF` / `COUNTIFS` · `SUMIF` / `AVERAGEIF` · `IFS` · Text functions (`MID`, `FIND`, `SUBSTITUTE`, `TRIM`) · `UNIQUE` / `SORT` / `LARGE` · `DATEVALUE` / `WEEKDAY` / `TEXT` · Pivot Tables · Charts & Histograms · Conditional Formatting · `FREQUENCY`

## Files

- `Zalando_Men_s_T-Shirt_Market_Analysis.xlsx` — full analysis workbook

---
*Dataset: Zalando.de men's t-shirt & polo catalogue, scraped January 2023. Practice/portfolio project — synthetic scrape recreation, not affiliated with Zalando.*
