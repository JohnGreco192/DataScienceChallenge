# E-commerce Data Analysis and Reporting in R

*Note: This project was originally completed in 2019. It is maintained here as a demonstration of long-standing proficiency in R and fundamental data analysis skills.*

## Project Overview

This script processes raw e-commerce data to calculate key metrics, identify trends, and produce actionable reports. It showcases data cleaning, aggregation, time-series analysis (Month-over-Month comparisons), and static data visualization.

## Key Features

* **Data Preprocessing:** Handles date formatting, feature engineering (`YearMonth`), and data cleaning.
* **Metric Calculation:** Computes Ecommerce Conversion Rate (ECR) for various segments.
* **Data Aggregation:** Summarizes data by month, device, and browser categories.
* **Comparative Analysis:** Performs Month-over-Month (MoM) comparisons for key performance indicators.
* **Reporting:** Generates a multi-sheet Excel file with aggregated data and MoM analysis.
* **Visualization:** Creates bar charts (Quantity by Device, Top Browsers) and a time-series line plot (Daily ECR) using `ggplot2`.

## Technologies Used

* **R Language**
* **R Libraries:** `tidyverse` (for `dplyr`, `ggplot2`), `zoo`, `DataExplorer`, `openxlsx`


## Output

The script will generate:
* An Excel file (`GrecoWB.xlsx`) in the project's root directory, containing aggregated data and MoM comparisons.
* Three plots displayed in the RStudio Plots pane: quantity by device, top browsers by quantity, and daily ECR time series.

## Data Source

Sample e-commerce data provided by Google.