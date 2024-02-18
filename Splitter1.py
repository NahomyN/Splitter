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
    and inserting them into new rows right below the original row. The content
    in the column immediately to the right of the steps column will be aligned
    with the last step, while the first step will be in the original row and
    all other columns maintain their original values, except for the column to
    the right of the steps column which should be empty in all but the last step row.

    :param dataframe: The original DataFrame
    :param step_column_index: The index of the column containing the steps
    :return: The reorganized DataFrame
    """
    new_rows = []
    for index, row in dataframe.iterrows():
        steps = split_cell_into_steps(row[step_column_index])
        if steps:
            # For the first step, keep all columns up to the step column with only the first step,
            # and empty the column immediately to the right of the steps column, keep the rest as is
            first_step_row = row.tolist()[:step_column_index] + [steps[0]] + [''] + row.tolist()[step_column_index + 2:]
            new_rows.append(first_step_row)

            # For intermediate steps, only fill the step column, keep others empty
            for step in steps[1:-1]:
                intermediate_row = [''] * step_column_index + [step] + [''] * (len(row) - step_column_index - 1)
                new_rows.append(intermediate_row)

            # For the last step, fill the step column and the column immediately to the right, keep others as is
            last_step_row = [''] * step_column_index + [steps[-1], row[step_column_index + 1]] + row.tolist()[step_column_index + 2:]
            new_rows.append(last_step_row)
        else:
            # Append the original row as it is
            new_rows.append(row.tolist())

    # Create a new DataFrame from the new_rows list, ensuring the correct number of columns
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
    #  input_file_path = 'C:\\Users\\YBira\\Documents\\Test Manangement Tools\\UAT_Regression_1.74 - Copy.xlsx'
    #  output_file_path = 'C:\\Users\\YBira\\Documents\\Test Manangement Tools\\UAT_Regression_1.74 - Copy - Splitted.xlsx'
     input_file_path = input('Enter input file path here:\n')#'C:\\Users\\YBira\\Documents\\Test Manangement Tools\\1.77.xlsx'
     output_file_path = input('Enter output file path here:\n')#'C:\\Users\\YBira\\Documents\\Test Manangement Tools\\1.77.xlsx-Splitted.xlsx'
     step_column_index = int(input("What is the index of the column you want to be splitted? \n Please enter the number: "))

     data = read_excel_file(input_file_path)
     if data is not None:
         reorganized_data = reorganize_data(data, step_column_index)
         write_to_excel(reorganized_data, output_file_path)


