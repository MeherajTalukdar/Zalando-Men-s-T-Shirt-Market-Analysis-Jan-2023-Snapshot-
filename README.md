# 🛍️ Zalando DE Product — Excel Practice Project

> **Dataset:** German Zalando men's basic t-shirts • **999 products** • **21 columns** • Scraped January 2023  


# Key Stats

Total products       999
Unique brands        145
Products on discount  43%  (432 products)
Average discount      33%
Price range          €6.99 – €99.99
Scan date            2023/1/29 01:31 (Sunday)


---

# 20 Excel Practice Questions

# Section 1 — Pricing & Discount Analysis

**Q1 — Discount Percentage** `formula`  
Calculate the discount percentage for each product using `PRICE_RETAIL` and `PRICE_CURRENT`. Which product has the highest markdown?

```excel
= (PRICE_RETAIL - PRICE_CURRENT) / PRICE_RETAIL
```

---

**Q2 — Count Discounted Products** `COUNTIF`  
How many products are currently on sale (current price below retail price)?

```excel
= COUNTIF(PRICE_CURRENT, "<" & PRICE_RETAIL)
```

---

**Q3 — Price Stats per Brand** `Pivot Table`  
What is the average, minimum, and maximum current price per brand?  
→ *Insert > PivotTable → Rows: BRAND → Values: PRICE_CURRENT (Average, Min, Max)*

---

**Q4 — Discount Tier Classification** `IFS`  
Classify each product into a discount tier:

| Tier | Condition |
|---|---|
| No Discount | 0% |
| Light | < 20% |
| Moderate | 20% – 50% |
| Deep | > 50% |

```excel
= IFS(disc=0, "No Discount", disc<0.2, "Light", disc<0.5, "Moderate", TRUE, "Deep")
```

---

**Q5 — Top 10 Biggest Price Drops** `LARGE` / `SORT`  
Which 10 products have the largest absolute price drop (retail minus current)?

```excel
= SORT(data, discount_col, -1)   ← then take top 10
= LARGE(discount_range, {1,2,...,10})
```

---

# Section 2 — Brand & Category Performance

**Q6 — Products per Brand** `COUNTIF`  
How many products does each brand have listed? Which brand has the most SKUs?

```excel
= COUNTIF($C$2:$C$1000, X2)
```

---

**Q7 — Brand Market Share** `COUNTIF` / `COUNTA`  
What is each brand's percentage share of the total product catalogue?

```excel
= COUNTIF($C$2:$C$1000, X2) / COUNTA($C$2:$C$1000)
```

---

**Q8 — Premium vs Budget Brands** `AVERAGEIF`  
Which brands position as premium vs. budget based on average retail price?

```excel
= AVERAGEIF($C$2:$C$1000, X2, $I$2:$I$1000)
```

---

**Q9 — Top 10 Brands Bar Chart** `Chart`  
Build a bar chart showing the top 10 brands by number of products listed.  
→ *Use Q6 result → select Brand + Count → Insert > Bar Chart → sort descending*

---

**Q10 — Revenue Potential per Brand** `SUMIF`  
What is the total revenue potential (sum of current prices) per brand?

```excel
= SUMIF($C$2:$C$1000, X2, $J$2:$J$1000)
```

---

# Section 3 — Product Lookup & Search

**Q11 — SKU Lookup Tool** `XLOOKUP`  
Create a lookup: enter a SKU and return the product name, brand, and current price.

```excel
= XLOOKUP(lookup_sku, $A$2:$A$1000, $H$2:$H$1000, "Not found")
```

---

**Q12 — Find "basic" Products** `SEARCH` + `IF`  
Flag all products whose name contains the word "basic". How many are there?

```excel
= IF(ISNUMBER(SEARCH("basic", H2)), "Yes", "No")
```

---

**Q13 — Extract Colour from Name** `MID` / `FIND` / `RIGHT`  
Extract the colour description from `PRODUCT_NAME` (text after the last dash `-`).

```excel
= TRIM(RIGHT(SUBSTITUTE(H2, "-", REPT(" ", 100)), 100))
```

---

**Q14 — Unique Sorted Brand List** `UNIQUE` / `SORT`  
Generate a unique, alphabetically sorted list of all brands in the catalogue.

```excel
= SORT(UNIQUE(C2:C1000))
```

---

**Q15 — Find Rank #1 Product** `INDEX-MATCH`  
Which product holds rank #1 in the sort ranking, and which brand does it belong to?

```excel
= INDEX(H2:H1000, MATCH(1, P2:P1000, 0))   ← product name
= INDEX(C2:C1000, MATCH(1, P2:P1000, 0))   ← brand
```

---

# Section 4 — Pivot Tables & Dashboards

**Q16 — Subcategory Breakdown** `Pivot Table`  
Build a pivot showing product count and average price by `SUBCATEGORY` and `SUBSUBCATEGORY`.  
→ *Rows: SUBCATEGORY + SUBSUBCATEGORY → Values: Count of SKU, Average of PRICE_CURRENT*

---

**Q17 — Price Distribution Histogram** `FREQUENCY` / Histogram  
How many products fall into each price band: €0–15, €15–30, €30–50, €50+?

```excel
= FREQUENCY(J2:J1000, {15, 30, 50})
```
→ *Or use Insert > Chart > Histogram directly on the PRICE_CURRENT column*

---

**Q18 — Highlight Heavy Discounts** `Conditional Formatting`  
Apply conditional formatting to highlight all rows where the discount exceeds 40%.  
→ *Select all data → Home > Conditional Formatting > New Rule > Use a formula:*

```excel
= ($I2 - $J2) / $I2 > 0.4
```

---

**Q19 — Heavily Discounted Products per Brand** `COUNTIFS`  
For each brand, how many products have a discount greater than 30%?

```excel
= COUNTIFS($C$2:$C$1000, X2, $V$2:$V$1000, ">0.3")
```
> ⚠️ Use `>0.3` not `>30` — Excel stores percentages as decimals.

---

**Q20 — Parse SCAN_DATE & Day of Week** `DATEVALUE` / `TEXT` / `WEEKDAY`  
Convert the raw text `2023/1/29 1:31` to a proper Excel date, then extract the day of the week.

```excel
= DATEVALUE(SUBSTITUTE(LEFT(A2, FIND(" ", A2)-1), "/", "-"))   ← convert to date
= WEEKDAY(W2, 2)                                                ← day number (Mon=1)
= TEXT(W2, "dddd")                                              ← day name
```
> 💡 All 999 rows share the same scan date: **Sunday, 29 January 2023** — the data was collected in a single overnight crawl.

---

# Skills Reference

| Skill | Questions |
|---|---|
| Arithmetic formulas | Q1, Q5, Q7 |
| COUNTIF / COUNTIFS | Q2, Q6, Q19 |
| SUMIF / AVERAGEIF | Q8, Q10 |
| IFS / IF / ISNUMBER | Q4, Q12 |
| XLOOKUP / INDEX-MATCH | Q11, Q15 |
| Text functions (LEFT, MID, FIND, SUBSTITUTE) | Q13, Q20 |
| UNIQUE / SORT / LARGE | Q5, Q14 |
| DATEVALUE / WEEKDAY / TEXT | Q20 |
| Pivot Tables | Q3, Q9, Q16 |
| Charts & Histograms | Q9, Q17 |
| Conditional Formatting | Q18 |
| FREQUENCY | Q17 |

---

#  Business Problems to Explore

| # | Problem | Columns Involved |
|---|---|---|
| 1 | Heavy discounting erodes margin — 43% of products are marked down avg. 33% | `PRICE_RETAIL`, `PRICE_CURRENT` |
| 2 | Long tail of low-visibility brands — top 3 own 23% of catalogue | `BRAND`, `SKU` |
| 3 | Sort ranking shows no correlation with price or discount | `SORT_BY_RANKING`, `PRICE_CURRENT` |
| 4 | Category skews heavily budget — 54% of products under €15 | `PRICE_CURRENT` |
| 5 | Price floor competition — same subcategory spans €6.99 to €99.99 | `BRAND`, `PRICE_CURRENT`, `SUBCATEGORY` |
| 6 | Data quality gaps — `PROMOTION` and `SKU_VARIANT` are 100% null | `PROMOTION`, `SKU_VARIANT`, `PRICE_LABEL` |

---
*Dataset source: Zalando.de product catalogue scrape · January 2023 · Men's T-Shirts & Polos category*
