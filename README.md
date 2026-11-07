# Data Cleaning and Transformation

## 1. Project Overview

This project focuses on applying data cleaning techniques, formatting tools, and data transformation operations using Microsoft Excel to prepare a raw product dataset for further analysis.

The primary objective is to develop practical skills in data preprocessing, standardization, and data quality improvement.

The assignment focuses on handling missing values, correcting inconsistencies, removing duplicate records, restructuring columns, and applying appropriate formatting rules to improve the accuracy, usability, and readability of the dataset.

Rather than directly analyzing raw data, this project demonstrates how to identify and resolve common data quality issues before performing meaningful analysis.

## 2. Project Objectives

- Identify and handle missing values using suitable techniques.
- Standardize inconsistent text formats in product names.
- Identify and correct spelling errors in product categories.
- Detect and remove duplicate records.
- Split and restructure product identification information.
- Combine multiple text columns into a single column.
- Apply appropriate currency and date formats.
- Use conditional formatting to improve data visualization.

## 3. Dataset Description

The project uses a product dataset containing the following columns:

| Column Name | Description |
|---|---|
| Product ID | Product identification containing manufacturing date and country code |
| Product Name | Name or description of the product |
| Brand Name | Brand associated with the product |
| Quantity | Available quantity of the product |
| Category | Product classification |
| Price | Product price |

The raw dataset was examined to identify missing values, inconsistent text entries, spelling mistakes, duplicate records, and formatting issues.

## 4. Data Cleaning and Preparation Process

### 4.1 Handling Missing Values

**Problem Identified:**

Some records contained missing values in the Price and Category columns. Missing values can affect calculations, summaries, and further data analysis.

**A. Price – Median Imputation**

**Method Used:** Median imputation.

The median value was calculated from the available Price values and used to handle missing Price entries. An additional column named `Imputed Price` was created to maintain the processed values.

**Formula:**

```excel
=MEDIAN(F2:F100)
```

This calculates the median of the Price column.

To replace blank values in a separate column:

```excel
=IF(F2="",MEDIAN($F$2:$F$100),F2)
```

**Why Median Imputation?**

Median imputation is useful when the dataset contains unusually high or low prices. Unlike the average, the median is less affected by extreme values.

**B. Category – Missing Value Handling**

**Method Used:** Unknown classification and product-based classification.

- Missing Category values were initially identified and marked as `Unknown`.
- A separate column named `Product-Based Classification` was created.
- Categories were assigned based on the corresponding Product Names.

**Why?**

Product names can provide useful information for identifying the appropriate category. Creating a separate classification column also helps retain the original category information.

### 4.2 Standardizing Product Names and Correcting Category Errors

**Problem Identified:**

Product names contained inconsistent capitalization, while the Category column contained spelling variations such as `Electronic` and `Electronicss`.

**A. Product Name Standardization**

**Method Used:** PROPER function.

A new column named `Updated Product Name` was created to standardize the capitalization of product names.

**Formula:**

```excel
=PROPER(B2)
```

**Example:**

| Original Product Name | Updated Product Name |
|---|---|
| laptop bag | Laptop Bag |
| wireless headphones | Wireless Headphones |
| smart watch | Smart Watch |

**Why?**

Consistent capitalization improves readability and makes product names more standardized for reporting and analysis.

**B. Category Spelling Correction**

**Method Used:** Sort and Find & Replace.

1. Selected the Category column.
2. Sorted the values from A to Z to identify inconsistent spellings.
3. Identified `Electronic` and `Electronicss` as incorrect entries.
4. Used Find and Replace (`Ctrl + H`) to replace both values with `Electronics`.

**Why?**

Different spellings of the same category may be treated as separate groups during analysis. Standardizing them ensures consistent category classification.

### 4.3 Identifying and Removing Duplicate Records

**Problem Identified:**

The dataset was checked for duplicate records that could affect the accuracy of the data and lead to incorrect counts or summaries.

**Method Used:** Remove Duplicates feature.

**Steps performed:**

1. Selected the complete dataset.
2. Navigated to the Data tab.
3. Selected Remove Duplicates.
4. Selected all columns to check duplicates based on the entire row.
5. Removed duplicate records identified by Excel.

**Why?**

Duplicate records can cause overcounting and affect calculations and reporting. Removing exact duplicate rows improves data consistency.

### 4.4 Splitting and Restructuring Product ID

**Problem Identified:**

The Product ID column contained multiple pieces of information in a single field, including the manufacturing date and country code.

**Example:**

`21-AUG-CA`

The information was separated into two columns:

- Manufacturing Date: `21-AUG`
- Country Code: `CA`

**Method Used:** TEXTBEFORE and TEXTAFTER functions.

**Manufacturing Date formula:**

```excel
=TEXTBEFORE(A2,"-", -1)
```

**Country Code formula:**

```excel
=TEXTAFTER(A2,"-", -1)
```

**Why?**

Separating these values makes the dataset more structured and allows manufacturing dates and country codes to be used independently for filtering, sorting, and analysis.

### 4.5 Converting Manufacturing Date into Date Format

**Problem Identified:**

The extracted Manufacturing Date was initially available as text, for example `21-AUG`. It needed to be converted into a proper date value and displayed in `DD-MM-YYYY` format.

**Method Used:** DATEVALUE with date formatting.

**Formula example:**

```excel
=DATE(2026,MONTH(DATEVALUE("1-"&MID(G2,4,3))),VALUE(LEFT(G2,2)))
```

Here, `G2` represents the extracted Manufacturing Date. The year 2026 is used as an example and should be replaced with the correct year if available.

**Steps performed:**

1. Extracted the day and month from the Manufacturing Date.
2. Used DATEVALUE to interpret the month.
3. Created a valid Excel date using the DATE function.
4. Applied Custom Date Formatting using `dd-mm-yyyy`.

**Why?**

Excel recognizes proper date values more effectively than text dates. This allows users to sort dates chronologically, filter by date, and perform date-based calculations.

### 4.6 Merging Brand Name and Product Name

**Problem Identified:**

Brand Name and Product Name were stored in separate columns. A combined field was useful for displaying complete product descriptions.

**Method Used:** TEXTJOIN function.

A new column named `Product Brand` was created by combining Brand Name and Product Name.

**Formula:**

```excel
=TEXTJOIN(" ",TRUE,C2,B2)
```

**Example:**

| Brand Name | Product Name | Product Brand |
|---|---|---|
| Samsonite | Laptop Bag | Samsonite Laptop Bag |
| Casio | Watch | Casio Watch |
| Sony | Headphones | Sony Headphones |

**Why?**

Combining related text fields creates a more descriptive product label, which is useful for reports and data presentation.

### 4.7 Applying Currency and Date Formatting

**A. Price Formatting**

**Method Used:** Currency format.

1. Selected the Price column.
2. Navigated to the Home tab.
3. Selected the Currency format from the Number group.

**Why?**

Currency formatting improves readability by displaying price values in a consistent monetary format.

**B. Manufacturing Date Formatting**

Applied the custom date format:

```text
dd-mm-yyyy
```

**Why?**

A consistent date format improves readability and avoids confusion when interpreting date values.

### 4.8 Conditional Formatting and Data Visualization

Conditional formatting was applied to highlight important information and make the dataset easier to understand.

**A. Data Bars – Price Column**

**Method Used:** Data Bars.

1. Selected the Price column.
2. Opened Conditional Formatting from the Home tab.
3. Selected Data Bars.
4. Applied a suitable bar style.

**Purpose:**

Data Bars provide a visual comparison of price values within the dataset. Higher values display longer bars, while lower values display shorter bars.

**B. Highlighting Electronics Category**

**Method Used:** New Rule in Conditional Formatting.

1. Selected the Category column.
2. Opened Conditional Formatting.
3. Selected New Rule.
4. Created a rule to highlight cells containing `Electronics`.
5. Applied a fill colour to identify matching categories.

**Purpose:**

Highlighting a specific category makes it easier to identify and filter important product groups.

## 5. Excel Functions and Features Used

| Function / Feature | Purpose |
|---|---|
| MEDIAN | Calculate the median price for missing value imputation |
| IF | Check for blank values and apply conditions |
| PROPER | Standardize product name capitalization |
| TEXTBEFORE | Extract manufacturing date from Product ID |
| TEXTAFTER | Extract country code from Product ID |
| DATEVALUE | Convert date text into an Excel date value |
| DATE | Construct a valid date |
| MID | Extract the month portion from the manufacturing date |
| LEFT | Extract the day portion from the manufacturing date |
| TEXTJOIN | Combine Brand Name and Product Name |
| Find and Replace | Correct spelling inconsistencies |
| Remove Duplicates | Remove duplicate records |
| Currency Formatting | Format price values |
| Conditional Formatting | Highlight specific values and patterns |
| Data Bars | Visually compare price values |

## 6. Final Outcome

The product dataset was cleaned, standardized, and restructured using Microsoft Excel.

The completed work improved the consistency and readability of the dataset through missing value handling, text correction, duplicate removal, column transformation, date and currency formatting, and conditional formatting.

The cleaned dataset is now better structured and more suitable for further analysis.

## 7. Key Learning Outcomes

Through this assignment, I gained practical experience in:

- Understanding common data quality issues.
- Selecting suitable methods for handling missing values.
- Standardizing text and correcting inconsistent entries.
- Identifying and removing duplicate records.
- Extracting and combining information using Excel functions.
- Converting text values into usable date formats.
- Applying formatting and visualization techniques.
- Preparing raw datasets for further analysis.

## 8. Project Deliverables

- Original raw product dataset.
- Cleaned and transformed Excel workbook.
- Assignment screenshots showing the cleaning process.
- Project documentation.

---

**Project Category:** Data Analytics – Excel

**Tool Used:** Microsoft Excel

**Focus Areas:** Data Cleaning | Data Preprocessing | Data Standardization | Data Transformation | Conditional Formatting
