12D Chain File Generator
This Python script reads data from an Excel file and generates chain files (.chain) for creating polygons in 12d Model software. The output files are segregated based on stages (Stage 2 and Stage 3) and are used to automate the creation of LOT polygons with specific attributes.YOU WILL NEED TO ADJUST THIS TO YOUR OWN PROJECT 

Table of Contents
Overview
Prerequisites
Installation
Usage
Input File Structure
Output Files
Customization
License
Overview
The script performs the following tasks:

Reads data from an Excel file (export.xlsx).
Filters and cleans the data based on specific criteria.
Iterates over the cleaned data to generate XML-formatted strings.

The generated chain files are intended to be used with 12d Model software to automate polygon creation.

Prerequisites
Python 3.x
pandas library
Installation
Clone the Repository

bash
Copy code
git clone https://github.com/kjarada/12D-Chain-File-Generator.git
cd chain-file-generator
Install Required Libraries


License
This project is licensed under the MIT License.

Disclaimer: This script is provided as-is and is intended for educational or developmental purposes. Always back up your data before running scripts that modify files.
