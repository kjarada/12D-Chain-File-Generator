#this is just an example of how to create a 12d Model chain using python, you need to adjust according to your project
#contact me if you need help


import pandas as pd
import chianTxtIntro
import ChainBodyXML

# Set pandas display option
pd.set_option("display.max_rows", None)

# Define the path to the Excel file
path_to_excel_file = 'export.xlsx'

# Alignment strings for Stage 2 and Stage 3 -- chainge this to your project alignment strings
alignment_stg2 = (
    r"<model_name>DES FRM STG2 Alignments V05</model_name>"
    r"<model_id>{3B53B647-02B1-4808-87C6-D0474E81EB80-0000000000002034}</model_id>"
    r"<name>N2NS MainLine 115 3 5</name>"
    r"<id>{3B53B647-02B1-4808-87C6-D0474E81EB80-0000000000002E46}</id>"
)

alignment_stg3 = (
    r"<model_name>DES FRM STG3 Alignments V05</model_name>"
    r"<model_id>{3B53B647-02B1-4808-87C6-D0474E81EB80-0000000000002034}</model_id>"
    r"<name>N2NS MainLine 115 3 5</name>"
    r"<id>{3B53B647-02B1-4808-87C6-D0474E81EB80-0000000000002E46}</id>"
)


def read_excel_file_and_return_data(path_to_excel_file):
    """
    Reads the Excel file and returns a cleaned DataFrame.
    Filters out rows based on specific criteria.
    """
    try:
        df = pd.read_excel(path_to_excel_file)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()  # Return empty DataFrame on error

    # Select relevant columns
    data = df[[
        "Lot No",
        "Chainage Start (km)",
        "Chainage End (km)",
        "Activity",
        "Region",
        "Description",
        "Status",
        "Sub Area"
    ]]

    # Apply filters
    data = data[
        (data['Chainage End (km)'] != 0) &
        (data['Chainage End (km)'] <= 1000) &
        (data['Chainage Start (km)'] <= 1000) &
        (data['Status'] != 'Abandoned')
    ]

    return data


def iter_over_data_and_write_chain(data, alignment_stg2, alignment_stg3, f_stg2, f_stg3):
    """
    Iterates over the data and writes the chain files for Stage 2 and Stage 3.
    """
    record_not_stored_counter = 0

    for index, row in data.iterrows():
        # Build the text using ChainBodyXML.Chaintext template
        name1 = f"{str(row['Lot No']).replace('/', '-')} {row['Description']} {row['Status']}"
        text = ChainBodyXML.Chaintext.replace("name1", name1)

        # Replace chainage start and end
        ch1 = str(row['Chainage Start (km)'] * 1000)
        ch2 = str(row['Chainage End (km)'] * 1000)
        text = text.replace("ch1", ch1)
        text = text.replace("ch2", ch2)

        # Determine the 'model1' string
        description_lower = str(row['Description']).lower()
        if "e3" in description_lower:
            if "sf" in description_lower:
                if "layer 1" in description_lower:
                    model1 = "PPW BDY Earthworks Filling E3 SF Layer 1"
                elif "layer 2" in description_lower:
                    model1 = "PPW BDY Earthworks Filling E3 SF Layer 2"
                else:
                    model1 = "PPW BDY Earthworks Filling E3 SF"
            else:
                if "layer 1" in description_lower:
                    model1 = "PPW BDY Earthworks Filling E3 Layer 1"
                elif "layer 2" in description_lower:
                    model1 = "PPW BDY Earthworks Filling E3 Layer 2"
                else:
                    model1 = "PPW BDY Earthworks Filling E3"
        elif "c3" in description_lower:
            if "layer 1" in description_lower:
                model1 = "PPW BDY Earthworks Filling C3 Layer 1"
            elif "layer 2" in description_lower:
                model1 = "PPW BDY Earthworks Filling C3 Layer 2"
            else:
                model1 = "PPW BDY Earthworks Filling C3"
        else:
            if "layer 1" in description_lower:
                model1 = "PPW BDY Earthworks Filling Layer 1"
            elif "layer 2" in description_lower:
                model1 = "PPW BDY Earthworks Filling Layer 2"
            else:
                model1 = f"PPW BDY {row['Activity']}"

        text = text.replace("model1", model1)

        # Determine color based on Status
        status = str(row['Status'])
        if status == "Open":
            color = "yellow"
        elif status == "Planned":
            color = "blue"
        elif status == "Closed":
            color = "green"
        else:
            color = "red"

        text = text.replace("colorr", color)

        # Determine the stage and write to the appropriate file
        region = str(row['Region'])
        if region.startswith("Stage 2"):
            text = text.replace("xxALxx", alignment_stg2)
            f_stg2.write(text)
        elif region.startswith("Stage 3"):
            text = text.replace("xxALxx", alignment_stg3)
            f_stg3.write(text)
        else:
            # Handle other stages if necessary
            record_not_stored_counter += 1
            print(f"Record at index {index} with Region '{region}' not stored.")

    print(f"Number of records not stored: {record_not_stored_counter}")


def main():
    # Read the data
    data = read_excel_file_and_return_data(path_to_excel_file)

    if data.empty:
        print("No data to process.")
        return

    # Open the chain files using context managers
    with open('PPW_STG2_Create_Polygons.chain', 'w') as f_stg2, open('PPW_STG3_Create_Polygons.chain', 'w') as f_stg3:
        # Write the introduction to both files
        f_stg2.write(chianTxtIntro.Intro)
        f_stg3.write(chianTxtIntro.Intro)

        # Process the data and write to files
        iter_over_data_and_write_chain(data, alignment_stg2, alignment_stg3, f_stg2, f_stg3)

        # Write the closing tags to both files
        closing_tags = '</Commands>\n</Chain>\n</xml12d>'
        f_stg2.write(closing_tags)
        f_stg3.write(closing_tags)

    print("Chain files written successfully.")


if __name__ == "__main__":
    main()
