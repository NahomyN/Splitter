import pandas as pd

# Function 1: Read Excel File
def read_excel_file(file_path):
    """
    Reads an Excel file and returns the data from the first sheet.
    :param file_path: The path to the Excel file (either .xlsx or .xls)
    :return: DataFrame containing the data from the first sheet
    """
    try:
        data = pd.read_excel(file_path)
        return data
    except Exception as e:
        print(f"An error occurred while reading the file: {e}")
        return None

# Function 2: Split Cell into Steps
def split_cell_into_steps(cell_content):
    """
    Splits the cell content into separate steps based on line breaks.
    :param cell_content: The text content of a cell
    :return: A list of steps if multiple steps are found, None otherwise
    """
    if pd.isna(cell_content) or cell_content.strip() == '':
        return None

    steps = cell_content.split('\n')
    if len(steps) > 1:
        return steps
    else:
        return None

# Function 3: Reorganize Data
def reorganize_data(dataframe, step_column_index):
    """
    Reorganizes the DataFrame by splitting the steps in the specified column
    and inserting them into new rows right below the original row.

    :param dataframe: The original DataFrame
    :param step_column_index: The index of the column containing the steps
    :return: The reorganized DataFrame
    """
    new_rows = []
    for index, row in dataframe.iterrows():
        steps = split_cell_into_steps(row[step_column_index])
        if steps:
            # Update the current row with the first step and append it
            new_row = row.tolist()[:step_column_index] + [steps[0]]
            new_rows.append(new_row)

            # Create and append new rows for additional steps
            for step in steps[1:]:
                empty_cols = [None] * step_column_index
                new_row = empty_cols + [step]
                new_rows.append(new_row)
        else:
            # Append the original row as it is
            new_rows.append(row.tolist())

    # Create a new DataFrame from the new_rows list
    new_dataframe = pd.DataFrame(new_rows, columns=dataframe.columns)
    return new_dataframe

# Example usage:
# reorganized_data = reorganize_data(dataframe, 4)
# This will reorganize the steps in the 5th column (since indexing starts at 0)



# Function 4: Write to Excel
def write_to_excel(dataframe, output_file_path):
    """
    Writes the given DataFrame to an Excel file.
    :param dataframe: The DataFrame to write to an Excel file
    :param output_file_path: The path where the Excel file will be saved
    """
    try:
        dataframe.to_excel(output_file_path, index=False)
        print(f"File successfully written to {output_file_path}")
    except Exception as e:
        print(f"An error occurred while writing the file: {e}")

# Main section (commented out for demonstration)
if __name__ == "__main__":
     input_file_path = 'C:\\Users\\yirna\\OneDrive\\Desktop\\split_row.xlsx'
     output_file_path = 'C:\\Users\\yirna\\OneDrive\\Desktop\\output.xlsx'
     
     data = read_excel_file(input_file_path)
     if data is not None:
         reorganized_data = reorganize_data(data)
         write_to_excel(reorganized_data, output_file_path)
