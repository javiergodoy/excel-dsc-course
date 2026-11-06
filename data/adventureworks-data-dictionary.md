# AdventureWorks Data Dictionary

Every chapter in this book works on the same dataset: the **AdventureWorks**
sample business, a fictional multinational manufacturer and retailer of
bicycles, components, clothing and accessories. This file is the single source
of truth for table names, grain and key columns. When a chapter says "open
`FactInternetSales`", this is the table it means.

## Why AdventureWorks

A consulting case needs a business with enough structure to be interesting and
enough mess to be realistic. AdventureWorks gives you:

- Two distinct sales channels with different economics (direct-to-consumer
  internet sales, and reseller/wholesale sales).
- Four product categories with very different margins and volumes.
- A customer dimension rich enough for real segmentation (income, education,
  commute distance, children at home, home ownership).
- Four calendar years of daily transactions, so trend and seasonality are
  genuinely visible rather than simulated.
- Geography down to city level across three sales groups (North America,
  Europe, Pacific).

That is enough to support a complete consulting engagement arc: diagnose, size,
segment, model, recommend.

## Which version to download

This book uses the **data-warehouse (star schema) flavour**, `AdventureWorksDW`.
It is the version shaped for analysis rather than for transaction processing:
facts in the middle, dimensions around the edges, surrogate integer keys
joining them.

Microsoft publishes it as a SQL Server backup and as CSV extracts. If you are
working in Excel only, you want the CSVs or the ready-made Excel workbook
version. Any of these will do:

- `AdventureWorksDW2019.bak` / `AdventureWorksDW2022.bak` restored to SQL Server
  or Azure SQL, then pulled into Excel with **Data → Get Data → From Database**.
- The CSV extracts of the DW tables, loaded with **Data → Get Data → From
  Text/CSV**.
- A pre-built `AdventureWorks.xlsx` containing the tables as worksheets.

Column names are stable across versions. Row counts differ slightly between
the 2012, 2017, 2019 and 2022 releases, so **never hard-code a row count or a
cell range in a formula**. The book uses structured table references
(`Sales[SalesAmount]`) throughout, which survive a version change.

## The star schema at a glance

```
                    DimDate
                       |
  DimCustomer --- FactInternetSales --- DimProduct --- DimProductSubcategory
       |               |                                        |
  DimGeography         |                               DimProductCategory
       |               |
 DimSalesTerritory ----+
                       |
  DimReseller --- FactResellerSales --- DimEmployee
                       |
                  DimPromotion
```

Read it as: **facts are events, dimensions are nouns.** Every number you will
ever sum lives in a fact table. Every label you will ever group by lives in a
dimension.

## Fact tables

### FactInternetSales

Direct-to-consumer orders placed on the AdventureWorks website.

**Grain: one row per product line on one sales order.** An order for three
different products is three rows. This is the single most important fact about
the table, and the source of most beginner errors.

| Column | Type | Notes |
|---|---|---|
| `ProductKey` | integer | → `DimProduct.ProductKey` |
| `OrderDateKey` | integer | → `DimDate.DateKey`, format `yyyymmdd` |
| `DueDateKey` | integer | → `DimDate.DateKey` |
| `ShipDateKey` | integer | → `DimDate.DateKey` |
| `CustomerKey` | integer | → `DimCustomer.CustomerKey` |
| `PromotionKey` | integer | → `DimPromotion.PromotionKey` |
| `CurrencyKey` | integer | → `DimCurrency.CurrencyKey` |
| `SalesTerritoryKey` | integer | → `DimSalesTerritory.SalesTerritoryKey` |
| `SalesOrderNumber` | text | e.g. `SO43697`; the business order id |
| `SalesOrderLineNumber` | integer | 1, 2, 3 … within the order |
| `OrderQuantity` | integer | units on this line; almost always 1 for internet sales |
| `UnitPrice` | currency | list price actually charged |
| `ExtendedAmount` | currency | `UnitPrice × OrderQuantity` |
| `UnitPriceDiscountPct` | decimal | 0 to 1 |
| `DiscountAmount` | currency | absolute discount on the line |
| `ProductStandardCost` | currency | cost of goods for one unit |
| `TotalProductCost` | currency | `ProductStandardCost × OrderQuantity` |
| `SalesAmount` | currency | **the revenue column; use this one** |
| `TaxAmt` | currency | tax |
| `Freight` | currency | shipping |
| `OrderDate`, `DueDate`, `ShipDate` | date | real dates, present in newer versions |

Roughly 60,000 rows covering 2010-12 to 2014-01 in the 2019+ releases.

**Gross margin is not in the table.** You derive it:
`SalesAmount − TotalProductCost`. Several chapters turn on this.

### FactResellerSales

Wholesale orders placed by bike shops and department stores.

**Grain: one row per product line on one reseller order.** Same shape as
internet sales, with three differences that matter commercially:

- `OrderQuantity` is frequently greater than 1 — resellers buy in bulk.
- `ResellerKey` and `EmployeeKey` replace `CustomerKey`: wholesale is a
  relationship sale with a named salesperson attached.
- Discounts are materially larger, so reseller revenue per unit is lower and
  mix effects are stronger.

| Column | Type | Notes |
|---|---|---|
| `ProductKey` | integer | → `DimProduct` |
| `OrderDateKey`, `DueDateKey`, `ShipDateKey` | integer | → `DimDate` |
| `ResellerKey` | integer | → `DimReseller` |
| `EmployeeKey` | integer | → `DimEmployee`, the owning salesperson |
| `PromotionKey` | integer | → `DimPromotion` |
| `CurrencyKey` | integer | → `DimCurrency` |
| `SalesTerritoryKey` | integer | → `DimSalesTerritory` |
| `SalesOrderNumber`, `SalesOrderLineNumber` | text, integer | order identity |
| `OrderQuantity` | integer | units; often 1–30 |
| `UnitPrice`, `ExtendedAmount` | currency | |
| `UnitPriceDiscountPct`, `DiscountAmount` | decimal, currency | |
| `ProductStandardCost`, `TotalProductCost` | currency | |
| `SalesAmount` | currency | revenue |
| `TaxAmt`, `Freight` | currency | |
| `CarrierTrackingNumber`, `CustomerPONumber` | text | mostly sparse |

Roughly 60,000 rows, but far more revenue than internet sales — around three
quarters of company turnover. Any analysis that silently uses only
`FactInternetSales` is wrong about the business.

### FactFinance

General-ledger style amounts by account, department, scenario and date. Used in
the financial-modelling chapter.

| Column | Type | Notes |
|---|---|---|
| `DateKey` | integer | → `DimDate` |
| `OrganizationKey` | integer | → `DimOrganization` |
| `DepartmentGroupKey` | integer | → `DimDepartmentGroup` |
| `ScenarioKey` | integer | 1 = Actual, 2 = Budget, 3 = Forecast |
| `AccountKey` | integer | → `DimAccount` |
| `Amount` | currency | signed; sign convention depends on account type |

## Dimension tables

### DimProduct

**Grain: one row per product, per version.** AdventureWorks tracks slowly
changing dimensions, so a product that changed price or cost can appear more
than once with different `ProductKey` values but the same
`ProductAlternateKey`. Counting distinct `ProductKey` overstates the catalogue.

| Column | Type | Notes |
|---|---|---|
| `ProductKey` | integer | surrogate key, joins to facts |
| `ProductAlternateKey` | text | business product code, e.g. `BK-R93R-62`; stable across versions |
| `ProductSubcategoryKey` | integer | → `DimProductSubcategory`; **null for non-sellable components** |
| `EnglishProductName` | text | e.g. `Road-150 Red, 62` |
| `StandardCost` | currency | cost at this version |
| `ListPrice` | currency | list price at this version |
| `Color`, `Size`, `Weight`, `Class`, `Style` | mixed | sparse; good cleaning practice |
| `DealerPrice` | currency | wholesale price |
| `StartDate`, `EndDate` | date | validity window of this version; `EndDate` null = current |
| `Status` | text | `Current` or null |

### DimProductSubcategory / DimProductCategory

The product hierarchy, in two hops.

- `DimProductSubcategory`: `ProductSubcategoryKey`, `EnglishProductSubcategoryName`
  (Mountain Bikes, Road Bikes, Touring Bikes, Helmets, Tires and Tubes, Jerseys …),
  `ProductCategoryKey`.
- `DimProductCategory`: `ProductCategoryKey`, `EnglishProductCategoryName` —
  exactly four values: **Bikes, Components, Clothing, Accessories**.

Category is the backbone of nearly every exhibit in this book. Bikes are high
price, low volume, thin percentage margin. Accessories are the opposite.

### DimCustomer

**Grain: one row per customer.** Internet sales only; resellers live in their
own dimension.

| Column | Type | Notes |
|---|---|---|
| `CustomerKey` | integer | surrogate key |
| `GeographyKey` | integer | → `DimGeography` |
| `CustomerAlternateKey` | text | business id, e.g. `AW00011000` |
| `FirstName`, `MiddleName`, `LastName` | text | name parts, not a full name |
| `BirthDate` | date | derive age from this, never store age |
| `MaritalStatus` | text | `M` / `S` |
| `Gender` | text | `M` / `F` |
| `YearlyIncome` | currency | banded in some versions |
| `TotalChildren`, `NumberChildrenAtHome` | integer | |
| `EnglishEducation` | text | Partial High School → Graduate Degree |
| `EnglishOccupation` | text | Clerical, Management, Manual, Professional, Skilled Manual |
| `HouseOwnerFlag` | text | `0` / `1` **stored as text** — a deliberate trap |
| `NumberCarsOwned` | integer | |
| `AddressLine1`, `AddressLine2` | text | `AddressLine2` mostly blank |
| `Phone` | text | inconsistent formatting; cleaning material |
| `DateFirstPurchase` | date | basis for cohort analysis |
| `CommuteDistance` | text | `0-1 Miles`, `1-2 Miles` … sorts wrong alphabetically |

About 18,500 customers.

### DimGeography

`GeographyKey`, `City`, `StateProvinceCode`, `StateProvinceName`,
`CountryRegionCode`, `EnglishCountryRegionName`, `PostalCode`,
`SalesTerritoryKey`.

Note that geography reaches territory, so there are two paths from a sale to a
country: through the customer's geography, and through the order's
`SalesTerritoryKey`. They do not always agree. Chapter 4 deals with this.

### DimSalesTerritory

`SalesTerritoryKey`, `SalesTerritoryRegion` (Northwest, Southwest, Central,
Northeast, Southeast, Canada, France, Germany, Australia, United Kingdom),
`SalesTerritoryCountry`, `SalesTerritoryGroup` (North America, Europe, Pacific).

There is an `NA` row with key 11 used for unassigned sales. Exclude it
deliberately, not accidentally.

### DimDate

**Grain: one row per calendar day.** A proper date dimension, and the reason
you will almost never write `=YEAR()` in this book.

| Column | Type | Notes |
|---|---|---|
| `DateKey` | integer | `yyyymmdd`, joins to every fact date |
| `FullDateAlternateKey` | date | the real date |
| `DayNumberOfWeek`, `EnglishDayNameOfWeek` | integer, text | |
| `DayNumberOfMonth`, `DayNumberOfYear` | integer | |
| `WeekNumberOfYear` | integer | |
| `EnglishMonthName`, `MonthNumberOfYear` | text, integer | |
| `CalendarQuarter`, `CalendarYear`, `CalendarSemester` | integer | |
| `FiscalQuarter`, `FiscalYear`, `FiscalSemester` | integer | fiscal year starts in July |

Calendar and fiscal years differ. A chart labelled "2013" means different things
to the controller and the sales director, and saying which you used is part of
the job.

### The FactFinance dimensions

`FactFinance` is served by four small dimensions that no other fact table uses.
They matter only in Chapter 13.

- **DimScenario**: `ScenarioKey`, `ScenarioName` — exactly three rows, `Actual`,
  `Budget`, `Forecast`. Any variance analysis on `FactFinance` is a comparison
  between two of these, and failing to filter the scenario double-counts the
  ledger.
- **DimAccount**: `AccountKey`, `ParentAccountKey` (a self-referencing chart of
  accounts hierarchy), `AccountCodeAlternateKey`, `AccountDescription`,
  `AccountType` (Revenue, Expenditures, Assets, Liabilities, Balances),
  `Operator` (`+` or `~`), `CustomMembers`. The `Operator` column is how the
  hierarchy signs its children; ignoring it inverts expense lines.
- **DimDepartmentGroup**: `DepartmentGroupKey`, `ParentDepartmentGroupKey`,
  `DepartmentGroupName` — Corporate, Executive General and Administration,
  Inventory Management, Manufacturing, Quality Assurance, Research and
  Development, Sales and Marketing.
- **DimOrganization**: `OrganizationKey`, `ParentOrganizationKey`,
  `CurrencyKey`, `OrganizationName` — the legal-entity structure used for
  consolidation.

### DimReseller

`ResellerKey`, `ResellerAlternateKey`, `GeographyKey`, `BusinessType`
(Warehouse, Value Added Reseller, Specialty Bike Shop), `ResellerName`,
`NumberEmployees`, `OrderFrequency`, `AnnualSales`, `AnnualRevenue`,
`YearOpened`. Around 700 resellers.

### DimEmployee

`EmployeeKey`, `ParentEmployeeKey` (the manager — a self-referencing hierarchy),
`FirstName`, `LastName`, `Title`, `HireDate`, `SalesTerritoryKey`,
`SalesPersonFlag`, `DepartmentName`, `Status`. Only rows with
`SalesPersonFlag = 1` appear in reseller sales.

### DimPromotion

`PromotionKey`, `EnglishPromotionName`, `DiscountPct`, `EnglishPromotionType`,
`EnglishPromotionCategory`, `StartDate`, `EndDate`, `MinQty`, `MaxQty`.

Key 1 is `No Discount` and covers most rows. Promotion analysis is therefore a
small-sample problem, which is itself a finding worth reporting.

### DimCurrency

`CurrencyKey`, `CurrencyAlternateKey` (ISO code), `CurrencyName`. Facts carry
`SalesAmount` already converted to US dollars, so currency is a label, not a
conversion you must perform.

## Known data quirks you will meet

These are not bugs in the book. They are the dataset being realistic, and
several exercises depend on them.

1. **`HouseOwnerFlag` is text `"0"`/`"1"`.** Summing it returns 0. Chapter 3.
2. **`DimProduct.ProductSubcategoryKey` is null** for manufacturing components
   that were never sold. Inner-join thinking silently drops them; outer-join
   thinking shows them as blank. Chapter 4.
3. **`CommuteDistance` sorts alphabetically**, putting `10+ Miles` before
   `2-5 Miles`. You need a custom sort order. Chapter 6.
4. **Products are versioned**, so `COUNTA(ProductKey)` is not "number of
   products". Chapter 2.
5. **Internet and reseller sales must be unioned** for a company view, and the
   two tables do not have identical columns. Chapter 7.
6. **The territory `NA` row** absorbs unassigned sales and will quietly appear
   in any unfiltered territory exhibit. Chapter 6.
7. **`DateKey` is an integer, not a date.** `20130715` minus `20130630` is not
   15 days. Chapter 2.
8. **Margin does not exist as a column.** Every margin number in the book is
   derived, which means every margin number can be derived differently by two
   analysts. Chapter 16 is about making that impossible.
9. **The last partial month** of the sales window looks like a collapse in
   revenue. It is a truncated period, not a trend. Chapter 10.
10. **Reseller `AnnualSales` in the dimension disagrees** with summed actual
    sales in the fact table. Dimension attributes are survey data; facts are
    transactions. Trust the facts and say so.

## Conventions used throughout this book

- Loaded tables are named without the `Dim`/`Fact` prefix where it reads better:
  `Sales`, `Resellers`, `Products`, `Customers`, `Calendar`.
- Formulas use **structured references** (`Sales[SalesAmount]`), never `A2:A60000`.
- Inputs are blue, formulas black, links to other sheets green — the standard
  financial-modelling convention, set out in Chapter 1.
- Currency is US dollars unless stated.
- "Revenue" always means `SalesAmount`. "Margin" always means
  `SalesAmount − TotalProductCost`. If a chapter needs a different definition it
  says so in its own `## Data you will use` section.
