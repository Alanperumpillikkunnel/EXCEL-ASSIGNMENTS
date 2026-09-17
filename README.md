# Excel Assignment Solutions

This repository contains my solutions and formulas for the Excel assignment, covering fundamental functions, conditional logic, text extraction, and data analysis.

## Assignment Questions & Answers

### 1. Total Price of All Products
* **Question:** What is the total price of all products in the dataset?
* **Formula Used:** `=SUM(D2:D35)`
* **Description:** Calculates the sum of all values in the price column.

### 2. Total Product Count
* **Question:** How many products are there in the dataset?
* **Formula Used:** `=COUNTA(B2:B35)`
* **Description:** Counts the number of non-empty cells in the product range.

### 3. Average Price
* **Question:** Calculate the average price of the products.
* **Formula Used:** `=AVERAGE(D2:D35)`
* **Description:** Finds the arithmetic mean of the product prices.

### 4. Minimum Price
* **Question:** Determine the minimum price among all products.
* **Formula Used:** `=MIN(D2:D35)`
* **Description:** Identifies the lowest price in the dataset.

### 5. Maximum Price
* **Question:** Find the maximum price among all products.
* **Formula Used:** `=MAX(D2:D35)`
* **Description:** Identifies the highest price in the dataset.

### 6. Price Range Category
* **Question:** Using an IF function, create a new column named `Price Range` to categorize products with a price $\ge$ $500 as 'High Price' and others as 'Standard Price'.
* **Formula Used:** `=IF(D2>=500, "High Price", "Standard Price")`
* **Description:** Evaluates each product's price conditionally and returns text labels.

### 7. Total Price for Electronics
* **Question:** Calculate the total price for products in the 'Electronics' category using the `SUMIF` function.
* **Formula Used:** `=SUMIF(F2:F35, "Electronics", D2:D35)`
* **Description:** Sums the prices specifically for rows where the category matches 'Electronics'.

### 8. Count of Products Under $100
* **Question:** Determine the count of products with a price less than $100 using the `COUNTIF` function.
* **Formula Used:** `=COUNTIF(D2:D35, "<100")`
* **Description:** Counts how many product prices meet the criteria of being under 100.

### 9. Extracting Day
* **Question:** Create a new column named `Day` with the first 2 characters of each 'Product ID' using the `LEFT` function.
* **Formula Used:** `=LEFT(A2, 2)`
* **Description:** Extracts a specified number of characters from the start of a text string.

### 10. Extracting Country Code
* **Question:** Create a new column named `Country Code` by extracting the last 2 characters from the 'Product ID' column using the `RIGHT` function.
* **Formula Used:** `=RIGHT(A2, 2)`
* **Description:** Extracts a specified number of characters from the end of a text string.

### 11. Extracting Month
* **Question:** Create a new column named `Month` by extracting 4th to 6th characters from the 'Product ID' column using the `MID` function.
* **Formula Used:** `=MID(A2, 4, 3)`
* **Description:** Extracts characters from the middle of a text string starting at a specific position for a given length.
