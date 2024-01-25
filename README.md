# Keyword Cannibalization Tool

This tool is designed to find pages that are competing with each other for the same keyword using your first-party Google Search Console (GSC) Data or SEMrush Keyword data. 

## Features

- **Data Source Check**: The tool checks the data source and validates whether it's from GSC or SEMrush based on the column names.
- **ASCII Check**: The tool removes rows containing non-ASCII characters.
- **Data Processing**: The tool processes the data by grouping by a specified metric, calculating sum per group and percentage per group, calculating cumulative sum and getting top n% of groups, merging sum_per_group and top_n_percent back to the original DataFrame, and more.
- **Data Merging**: The tool merges dataframes based on common columns.
- **Excel Formatting**: The tool formats the output Excel file by highlighting the first instance of each query (the best query/page match by traffic, and avg position). 

## Benefits

- **Efficiency**: This tool automates the process of finding keyword cannibalization, saving you time and effort.
- **Accuracy**: The tool uses precise calculations and data processing techniques to ensure accurate results.
- **Ease of Use**: The tool provides a user-friendly interface with clear instructions and error messages.
- **Customizability**: You can set a threshold to select the top n% of queries.

## How to Use

1. Set a threshold (e.g., 80 = Selecting the top 80% of queries by metric).
2. Upload your query + page data (.csv) from the GSC API or SEMrush Keyword Export.

The tool will process the data and provide an Excel file with the analysis results.

## Live App
You can run a live version of the [seo keyword cannibalization app here](https://keywordcannibalization.streamlit.app/).

## Author

[Tyler Gargula](https://tylergargula.dev) - Technical SEO & Software Developer