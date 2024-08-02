# Excel Macros Collection

This repository contains a collection of useful Excel macros for various tasks, such as data processing, file management, and analysis. Each macro is stored in its own Excel file and is designed to automate common tasks in Excel.

## List of Macros

### Data Processing

- **2_ProcessedSheet_to_GISTable.xlsm**: Aggregates GIS data from multiple workbooks into a single worksheet and saves the consolidated data into a new Excel file.
  - **Usage**:
    1. Open the workbook.
    2. Enable macros.
    3. Run the `CreateGISTable` macro.
    4. Select the input and output folders.
    5. View the results in the output folder.

- **convert_csv_to_xlsx.xlsm**: Converts CSV files to XLSX format.
  - **Usage**:
    1. Open the workbook.
    2. Enable macros.
    3. Run the `ConvertCSVToXLSX` macro.
    4. Select the CSV file and save location for the XLSX file.
    5. View the converted file.

- **selecting x rows.xlsm**: Selects and processes a specified number of rows from a dataset.
  - **Usage**:
    1. Open the workbook.
    2. Enable macros.
    3. Run the `SelectXRows` macro.
    4. Enter the number of rows to select.
    5. View the selected rows in the new worksheet.

### File Management

- **check if file exists.xlsm**: Checks if a specified file exists in a directory.
  - **Usage**:
    1. Open the workbook.
    2. Enable macros.
    3. Run the `CheckIfFileExists` macro.
    4. Select the file to check.
    5. View the result.

- **copy files_aa.xlsm**: Copies files from one directory to another.
  - **Usage**:
    1. Open the workbook.
    2. Enable macros.
    3. Run the `CopyFiles` macro.
    4. Select the source and destination folders.
    5. View the copied files in the destination folder.

- **Cut paste files .xlsm**: Moves files from one directory to another.
  - **Usage**:
    1. Open the workbook.
    2. Enable macros.
    3. Run the `CutPasteFiles` macro.
    4. Select the source and destination folders.
    5. View the moved files in the destination folder.

- **find external links.xlsm**: Finds and lists all external links within an Excel workbook.
  - **Usage**:
    1. Open the workbook.
    2. Enable macros.
    3. Run the `FindExternalLinks` macro.
    4. View the list of external links in the new worksheet.

- **finding contents of a folder.xlsm**: Lists the contents of a specified folder, including files and subfolders.
  - **Usage**:
    1. Open the workbook.
    2. Enable macros.
    3. Run the `ListFolderContents` macro.
    4. Select the folder to list contents for.
    5. View the list of contents in the new worksheet.

- **Folder and file size analysis.xlsm**: Analyzes the size of folders and files within a specified directory.
  - **Usage**:
    1. Open the workbook.
    2. Enable macros.
    3. Run the `AnalyzeFolderSize` macro.
    4. Select the folder to analyze.
    5. View the size analysis in the new worksheet.

### Analysis

- **copy specific cells from multiple files.xlsm**: Copies specific cells from multiple Excel files and consolidates them into a single file.
  - **Usage**:
    1. Open the workbook.
    2. Enable macros.
    3. Run the `CopySpecificCells` macro.
    4. Select the source folder containing the files.
    5. View the consolidated data in the new worksheet.

- **copying cells from specific files in folder.xlsm**: Copies specific cells from specified files within a folder and consolidates them into a single worksheet.
  - **Usage**:
    1. Open the workbook.
    2. Enable macros.
    3. Run the `CopyCellsFromSpecificFiles` macro.
    4. Select the source folder containing the files.
    5. View the consolidated data in the new worksheet.

- **Convert-Easting-Northing-to-Lattitude-Longitude-in-Excel.xlsm**: Converts Easting and Northing coordinates to Latitude and Longitude.
  - **Usage**:
    1. Open the workbook.
    2. Enable macros.
    3. Run the `ConvertEastingNorthing` macro.
    4. Enter the Easting, Northing, UTM Zone, and Hemisphere values.
    5. View the converted Latitude and Longitude in the specified cells.

- **MakeXLSXchanges_template.xlsm**: Applies predefined changes to an Excel file based on a template.
  - **Usage**:
    1. Open the workbook.
    2. Enable macros.
    3. Run the `ApplyTemplateChanges` macro.
    4. Select the source file to modify based on the template.
    5. View the modified file.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or new macros.

## License

This project is licensed under the Roy License. See the [LICENSE](LICENSE) file for details.

## Contact

For any questions or suggestions, feel free to contact the project maintainer.
