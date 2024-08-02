**Objective**:
A Python-based application with a graphical user interface (GUI) that allows users to compare multiple columns between two Excel files and generate a detailed comparison report. The tool aims to streamline the process of identifying matching and non-matching values across specified columns, thus improving data analysis efficiency for business users.

**Key Features**:

File Selection:
Prompt the user to select two Excel files to be compared.
Ensure flexibility in file formats (supporting .xlsx and .xls).

Column Selection:
Provide a scrollable, checkbox-based interface for users to select one or more columns from each Excel file.
Ensure selected columns from both files match in number.

Column Mapping:
Allow users to map each selected column from the first file to a corresponding column in the second file through an intuitive GUI.

Data Comparison:
Compare values in the mapped columns from both files.
Identify and categorize matching and non-matching values for each column pair.

Report Generation:
Generate an Excel report with detailed analysis:
Sheet 1 (Analysis): Summary of total values, matching values, and non-matching values for each column pair.
Sheet 2: Values present in the first file but not in the second.
Sheet 3: Values present in the second file but not in the first.
Ensure the report uses concise and valid worksheet names.

User-Friendly Saving:
Prompt the user to select the location and name for the report file.
Enforce the Excel file format for the output report.

Business Benefits:
Improved Data Accuracy: Helps in identifying discrepancies between two datasets, leading to more accurate and reliable data analysis.
Efficiency: Automates the comparison process, saving time and reducing manual effort.
User-Friendly: Easy-to-use interface allows business users with minimal technical expertise to perform complex data comparisons.
Flexibility: Supports multiple column comparisons and provides detailed insights, making it adaptable to various business scenarios.

By implementing this tool, businesses can ensure data consistency, improve analysis workflows, and make informed decisions based on accurate data comparisons.
