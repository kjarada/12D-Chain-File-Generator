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
Writes the generated strings into two separate chain files:
PPW_STG2_Create_Polygons.chain for Stage 2 alignments.
PPW_STG3_Create_Polygons.chain for Stage 3 alignments.
The generated chain files are intended to be used with 12d Model software to automate polygon creation.

Prerequisites
Python 3.x
pandas library
Installation
Clone the Repository

bash
Copy code
git clone https://github.com/yourusername/chain-file-generator.git
cd chain-file-generator
Install Required Libraries

Ensure you have pandas installed:

bash
Copy code
pip install pandas
Prepare Input Files

Place your export.xlsx file in the same directory as the script.
Ensure you have the chianTxtIntro.py and ChainBodyXML.py modules available, or adjust the import statements accordingly.
Usage
Run the script using the following command:

bash
Copy code
python chain_file_generator.py
Note: Replace chain_file_generator.py with the actual filename if it's different.

Input File Structure
The script expects an Excel file (export.xlsx) with the following columns:

Lot No
Chainage Start (km)
Chainage End (km)
Activity
Region
Description
Status
Sub Area
Data Filtering Criteria
The script filters out rows based on the following conditions:

Rows where Chainage End (km) is 0 or greater than 1000.
Rows where Chainage Start (km) is greater than 1000.
Rows where Status is Abandoned.
Output Files
The script generates two chain files:

Stage 2 Chain File

Filename: PPW_STG2_Create_Polygons.chain
Contains polygons related to Stage 2 alignments.
Stage 3 Chain File

Filename: PPW_STG3_Create_Polygons.chain
Contains polygons related to Stage 3 alignments.
Chain File Structure
Each chain file contains XML-formatted data that can be imported into 12d Model software. The polygons are created with attributes based on the input data, such as name, chainage, model, and color.

Customization
Adjusting Alignment Information
The alignment information is hardcoded in the script via the ALSTG2 and ALSTG3 variables. If your alignments are different, you can modify these variables:

python
Copy code
ALSTG2 = r"<model_name>Your Stage 2 Model Name</model_name><model_id>Your Model ID</model_id><name>Your Alignment Name</name><id>Your Alignment ID</id>"
ALSTG3 = r"<model_name>Your Stage 3 Model Name</model_name><model_id>Your Model ID</model_id><name>Your Alignment Name</name><id>Your Alignment ID</id>"
Modifying Output Paths
To change the output filenames or directories, modify the FSTG2 and FSTG3 variables:

python
Copy code
FSTG2 = open('Your_Custom_Path/Stage2.chain', "w")
FSTG3 = open('Your_Custom_Path/Stage3.chain', "w")
Changing Color Mapping
The color of the polygons is determined by the Status field. You can adjust the color mapping in the iter_over_data_and_write_chain function:

python
Copy code
if str(row[6]) == "Open":
    text = text.replace("colorr", "yellow")
elif str(row[6]) == "Planned":
    text = text.replace("colorr", "blue")
elif str(row[6]) == "Closed":
    text = text.replace("colorr", "green")
else:
    text = text.replace("colorr", "red")
Replace the color names with any valid color recognized by 12d Model.

License
This project is licensed under the MIT License.

Disclaimer: This script is provided as-is and is intended for educational or developmental purposes. Always back up your data before running scripts that modify files.