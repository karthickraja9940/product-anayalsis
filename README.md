# product-anaya# 📊 Data Analytics – Excel Assignment 2: Data Exploration

## 📌 Project Overview

This project is part of my **Data Analytics learning journey** and focuses on **data cleaning and preparation using Microsoft Excel**.

Real-world datasets often contain missing values, inconsistent formatting, duplicate records, spelling errors, and poorly structured data. In this project, I cleaned and transformed a product dataset to improve its quality and prepare it for further analysis.

### 🎯 Objective

The main objective of this assignment is to develop practical skills in:

* Data cleaning
* Data preprocessing
* Data standardization
* Data transformation
* Data formatting
* Data quality improvement
* Preparing datasets for analysis

---

## 📂 Dataset Information

**Dataset Name:** Product Dataset

### Attributes

| Column       | Description               |
| ------------ | ------------------------- |
| Product ID   | Unique product identifier |
| Product Name | Name of the product       |
| Brand Name   | Product brand             |
| Quantity     | Available quantity        |
| Category     | Product category          |
| Price        | Product price             |

---

## 🛠️ Tasks Performed

### 1. Handling Missing Values

* Checked the **Price** column for missing values.
* Identified appropriate approaches for products with missing prices.
* Checked the **Category** column for missing values.
* Considered suitable strategies for handling missing categories.

### 2. Correcting Inconsistent Data

* Identified inconsistent text formatting in the **Product Name** column.
* Identified spelling mistakes and inconsistencies in the **Category** column.
* Used **Find & Replace** to standardize product names.
* Corrected category typos and standardized category names.

### 3. Removing Duplicate Records

* Checked the complete dataset for duplicate rows.
* Removed duplicate records where applicable.
* Ensured each row represents a unique record.

### 4. Splitting and Merging Data

* Split the **Product ID** into:

  * Manufacturing Date
  * Country Code
* Removed unnecessary characters where required.
* Merged **Brand Name** and **Product Name** into a new column:

  * `Product Brand`

### 5. Number and Date Formatting

* Converted the **Price** column into currency format.
* Formatted the **Manufacturing Date** as:

`DD-MM-YYYY`

### 6. Conditional Formatting

* Applied **Data Bars / Color Scales** to the Price column.
* Created a custom conditional formatting rule to highlight:

`Electronics`

in the **Category** column.

---

## 🧰 Excel Skills Used

* Find & Replace
* Data Cleaning
* Data Validation
* Remove Duplicates
* Text Functions

  * `LEFT()`
  * `RIGHT()`
  * `MID()`
* IF / Conditional Logic
* Conditional Formatting
* Data Bars
* Color Scales
* Currency Formatting
* Date Formatting
* Splitting Columns
* Merging Columns

---

## 📈 Before and After

### Before Cleaning

The original dataset contained examples of:

* Inconsistent capitalization
* Category spelling inconsistencies
* Unstructured Product IDs
* Different text formats
* Potential duplicate records
* Data requiring better formatting

### After Cleaning

The cleaned dataset contains:

* Standardized product names
* Corrected category values
* Structured Manufacturing Date
* Extracted Country Code
* Combined Product Brand field
* Proper currency formatting
* Standardized date formatting
* Conditional formatting for easier analysis

---

## 📁 Repository Files

```text
Excel-Assignment-2-Data-Exploration/
│
├── Product_Dataset.xlsx
└── README.md
```

### `Product_Dataset.xlsx`

Contains the Excel dataset after applying the required data cleaning and preparation techniques.

### `README.md`

Contains the project documentation, objectives, tasks performed, and Excel skills demonstrated.

---

## 🎯 Outcome

This project helped me develop practical experience in **data cleaning and preparation using Microsoft Excel**.

The cleaned dataset is more structured, consistent, readable, and ready for further analysis and visualization.

### Key Learning

> **Clean Data → Better Analysis → Better Insights**

---

## 👨‍💻 About Me

I am an **Aspiring Data Analyst** currently developing my skills in:

* Microsoft Excel
* Data Cleaning
* Data Analysis
* SQL
* Data Visualization
* Python

This repository documents my progress and practical projects as I build my **Data Analytics portfolio**.

**Created by: Karthick Raja**
**Role: Aspiring Data Analyst**
lsis# 📊 Product Data Analysis Using Microsoft Excel

## 📌 Project Overview

This project is part of my journey to becoming a **Data Analyst**.

In this project, I performed basic data exploration and analysis using **Microsoft Excel**. The dataset contains information about different products, including Product IDs, Product Names, Brand Names, Quantities, Categories, and Prices.

The main objective of this project was to develop foundational skills in:

* Data exploration
* Data summarization
* Excel formulas and functions
* Conditional logic
* Conditional aggregation
* Text manipulation
* Data categorization

---

## 📂 Dataset

The dataset contains the following columns:

| Column       | Description                        |
| ------------ | ---------------------------------- |
| Product ID   | Unique identifier for each product |
| Product Name | Name of the product                |
| Brand Name   | Brand associated with the product  |
| Quantity     | Number of products                 |
| Category     | Product category                   |
| Price        | Product price                      |

---

## 🛠️ Tools & Technologies

* **Microsoft Excel**
* Excel Functions & Formulas
* Data Cleaning & Exploration
* GitHub

---

## 🔍 Analysis Performed

### 1. Basic Data Exploration

I used Excel functions to calculate:

* **Total Price** – using `SUM`
* **Number of Products** – using `COUNT`
* **Average Price** – using `AVERAGE`

Example formulas:

```excel
=SUM(F2:F100)
=COUNT(F2:F100)
=AVERAGE(F2:F100)
```

---

### 2. Minimum and Maximum Price

I identified the lowest and highest product prices using:

```excel
=MIN(F2:F100)
=MAX(F2:F100)
```

This helped understand the price range within the dataset.

---

### 3. Price Range Classification

I created a new column called **Price Range** using the `IF` function.

Business rule:

* Price ≥ $500 → **High Price**
* Price < $500 → **Standard Price**

Formula:

```excel
=IF(F2>=500,"High Price","Standard Price")
```

This categorizes products based on their price.

---

### 4. Conditional Analysis

#### Electronics Category

I used `SUMIF` to calculate the total price of products belonging to the **Electronics** category.

```excel
=SUMIF(E:E,"Electronics",F:F)
```

#### Products Below $100

I used `COUNTIF` to determine the number of products with a price below $100.

```excel
=COUNTIF(F:F,"<100")
```

---

## 🔤 5. Text Extraction from Product ID

I created three additional columns using Excel text functions.

### Day

Extracted the first two characters from Product ID using `LEFT`.

```excel
=LEFT(A2,2)
```

### Country Code

Extracted the last two characters from Product ID using `RIGHT`.

```excel
=RIGHT(A2,2)
```

### Month

Extracted characters 4 to 6 from Product ID using `MID`.

```excel
=MID(A2,4,3)
```

These functions helped me understand how information can be extracted from structured text fields.

---

## 📈 Key Skills Demonstrated

Through this project, I practiced:

* `SUM`
* `COUNT`
* `AVERAGE`
* `MIN`
* `MAX`
* `IF`
* `SUMIF`
* `COUNTIF`
* `LEFT`
* `RIGHT`
* `MID`
* Basic data exploration
* Data categorization
* Conditional analysis
* Text extraction

---

## 📁 Project Structure

```text
Product-Data-Analysis-Excel/
│
├── README.md
│
├── data/
│   └── product_dataset.xlsx
│
├── analysis/
│   └── product_analysis.xlsx
│
└── screenshots/
    ├── basic_analysis.png
    ├── price_range.png
    └── text_extraction.png
```

---

## 🎯 Learning Outcome

This project helped me build a foundation in **Excel-based data analysis** and understand how formulas can be used to summarize, categorize, and extract information from datasets.

I am continuing to develop my skills in:

* Microsoft Excel
* SQL
* Data Analysis
* Data Visualization
* Python
* Power BI

This project is one of the first projects in my **Data Analyst portfolio**.

---

## 👨‍💻 About Me

I am an aspiring **Data Analyst** currently building my skills through practical projects and hands-on learning.

I am interested in using data to identify patterns, generate insights, and support better decision-making.

### 🚀 My Data Analyst Learning Journey

**Excel → SQL → Power BI → Python → Data Analytics Projects**

---

## ⭐ Future Improvements

I plan to improve this project by adding:

* Excel Pivot Tables
* Charts and dashboards
* More detailed data cleaning
* SQL analysis
* Power BI visualization
* Python-based analysis

---

## 📌 Project Status

**Completed – Beginner Excel Data Analysis Project**
