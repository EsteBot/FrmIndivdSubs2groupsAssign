''' pysimpleGUI for analyzing behavior data and creating all possible group assignments.
    Total subjects per group assignment  have equal or near equal subjects within them. The optimal group 
    assignment out of all possible, is determined by finding mean values from hypothetical possible group 
    assignments for each columns data. The group assigment that results in the minimum value when the sum of 
    the mean's between groups is calculated, is assigned as the 'optimal' group. Each step of these 
    calculations create a new data frame that is saved to an Excel sheet along with the original data for 
    reference. The optimal group assigment is graphed for comparision.  
'''

# Import required libraries
import PySimpleGUI as sg
import pandas as pd
import itertools
import re
import matplotlib.pyplot as plt
import os
from pathlib import Path

# validate that the file paths are entered correctly
def is_valid_path(filepath):
    if filepath and Path(filepath).exists():
        return True
    sg.popup_error("A selected file path is incorrect or has been left empty.")
    return False

# window appears when the program successfully completes
def nom_window(input_filename):
    layout = [[sg.Text("\n"
    " All Systems Nominal.\n\n"
    f" A file: {input_filename} \n"
    " has been created\n"
    " containing info & calcs for\n"
    " optimal group assignments."
    "\n"
    "")]]
    window = sg.Window((""), layout, modal=True)
    choice = None
    while True:
        event, values = window.read()
        if event == "Exit" or event == sg.WIN_CLOSED:
            break
    window.close()
    
# Define the location of the directory
def extract_values_from_excel(input_filename, output_folder):
    name = Path(output_folder)

    # Change the directory
    os.chdir(output_folder)
    print(output_folder)

    # creation of a maximum value for the progress bar function
    max = 6
    prog_bar_update_val = 0

    file_name = input_filename

    # Extract the name before the extension
    base_name, extension = os.path.splitext(file_name)

    # Modify the base name (customize as needed)
    modified_base_name = base_name.replace(base_name, "eBot_Groups")  # Example modification

    # Create the new file name with the modified base name and .xlsx extension
    new_file_name = modified_base_name + extension 


    # Read the Excel file (assuming it's named 'data.xlsx')
    df = pd.read_excel(file_name)

    # Display the initial DataFrame
    #print("Initial DataFrame:")
    #print(df)

    df.to_excel(new_file_name, index=False)

    prog_bar_update_val += 1
    # records progress by updating prog bar with each file compiled
    window["-Progress_BAR-"].update(max = max, current_count=int(prog_bar_update_val))


    ###################### Generate all possible Group Assignments ######################

    # Extract column names except for 'group' and 'comboID'
    data_columns = [col for col in df.columns if col not in ['group', 'combo', 'id']]

    # Generate all possible combinations of 'a' and 'b' assignments
    num_ids = len(df)
    half_num_ids = num_ids // 2

    # Function to generate combinations with equal counts of 'a' and 'b'
    def generate_combinations(num_ids):
        for r in range(half_num_ids + 1):
            for combination in itertools.combinations(range(num_ids), r):
                yield ['a' if i in combination else 'b' for i in range(num_ids)]

    all_group_combinations = list(generate_combinations(num_ids))

    # Create a new DataFrame to store mean values for each combination
    mean_values_df = pd.DataFrame(columns=['group'] + data_columns)

    # Create a new DataFrame to store valid combinations
    all_combos_df = pd.DataFrame()

    # Counter for iteration number
    iteration_count = 1

    # Iterate over all combinations and filter those with an equal number of 'a' and 'b'
    for combination in all_group_combinations:
        a_count = combination.count('a')
        b_count = combination.count('b')
        
        if abs(a_count - b_count) <= 1:  # Allowing difference of 1 for uneven number of IDs
            temp_df = df.copy()
            temp_df['group'] = combination
            temp_df['combo'] = iteration_count
            all_combos_df = pd.concat([all_combos_df, temp_df], ignore_index=True)
            iteration_count += 1

            # Calculate mean values for each data column for 'a' and 'b'
            mean_values = temp_df.groupby('group')[data_columns].mean().reset_index()
            
            # Collect id numbers for each 'a' and 'b' group
            subj_ids = temp_df.groupby('group')['id'].apply(list).reset_index()
            
            # Merge ids with mean_values
            mean_values = pd.merge(mean_values, subj_ids, on='group')
            
            # Append the mean values to the DataFrame
            mean_values_df = pd.concat([mean_values_df, mean_values], ignore_index=True)

    #print(all_combos_df)
    # Save the updated DataFrame to a new sheet in the same Excel file
    with pd.ExcelWriter(new_file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        all_combos_df.to_excel(writer, sheet_name='AllPosCombos', index=False)

    #print(mean_values_df)
    # Save the updated DataFrame to a new sheet in the same Excel file
    with pd.ExcelWriter(new_file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        mean_values_df.to_excel(writer, sheet_name='AllPosMeans', index=False)

    prog_bar_update_val += 1
    # records progress by updating prog bar with each file compiled
    window["-Progress_BAR-"].update(max = max, current_count=int(prog_bar_update_val))

    ###################### Abs Diff from Calculated Mean Vals ######################

    # Read data from Excel file
    df = mean_values_df

    # Create a new dataframe for absolute differences
    abs_diff_df = pd.DataFrame()

    # Extract time point columns
    time_point_columns = [col for col in df.columns if col not in ['group', 'comboID', 'id']]

    # Iterate through time points
    for i in range(len(time_point_columns)):
        # Extract 'a' and 'b' values for the current time point
        a_values = df[df['group'] == 'a'][time_point_columns[i]].values
        b_values = df[df['group'] == 'b'][time_point_columns[i]].values

        # Extract time point number from column name using regular expressions
        time_point = i

        # Calculate the absolute difference and create a new column in the new dataframe
        abs_diff_df[f'AbsDiff_{time_point}'] = abs(a_values - b_values)

    # Collect ids and groups used for this time point
        ids_a = df[df['group'] == 'a']['id'].values
        ids_b = df[df['group'] == 'b']['id'].values

    # Create a list of ids with corresponding groups
        ids_groups = [f"{group_a}: {id_a}, {group_b}: {id_b}" for group_a, id_a, group_b, id_b in zip(['a']*len(ids_a), ids_a, ['b']*len(ids_b), ids_b)]

    # Add the 'idsNgroups' column to abs_diff_df
    abs_diff_df['Groups_id'] = ids_groups

    # Display the new dataframe
    #print("Absolute Mean Value Differences")
    #print(abs_diff_df)

    # Save the updated DataFrame to a new sheet in the same Excel file
    with pd.ExcelWriter(new_file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        abs_diff_df.to_excel(writer, sheet_name='MeanValsAbsDiff', index=False)

    prog_bar_update_val += 1
    # records progress by updating prog bar with each file compiled
    window["-Progress_BAR-"].update(max = max, current_count=int(prog_bar_update_val))

    ###################### Min Groups Diff from Abs Diff Mean Vals ######################

    # Sum the values of each row
    row_sums = abs_diff_df.iloc[:, :-1].sum(axis=1)

    # Find the minimum sum value(s) and their index(es)
    min_sum = row_sums.min()
    min_indices = row_sums[row_sums == min_sum].index

    # Create a new DataFrame to hold the lowest sum values and their 'Groups_id' information
    min_sum_values_df = pd.DataFrame(columns=['MinSumValue', 'Groups_id'])

    # Populate the new DataFrame with the lowest sum values and their 'idsNgroups' information
    for idx in min_indices:
        new_row = pd.DataFrame({
            'MinSumValue': [min_sum],
            'Groups_id': [abs_diff_df.loc[idx, 'Groups_id']]
        })
        min_sum_values_df = pd.concat([min_sum_values_df, new_row], ignore_index=True)

    # Display the new DataFrame
    #print(min_sum_values_df)

    # Save the updated DataFrame to a new sheet in the same Excel file
    with pd.ExcelWriter(new_file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        min_sum_values_df.to_excel(writer, sheet_name='MinMeansIDsGroups', index=False)

    prog_bar_update_val += 1
    # records progress by updating prog bar with each file compiled
    window["-Progress_BAR-"].update(max = max, current_count=int(prog_bar_update_val))

    ###################### Extract 'a', 'b' group assignments from list and create dictionary ######################
        
    # List of assigned groups
    string_list_element = [abs_diff_df.loc[idx, 'Groups_id']]

    print(string_list_element)

    # Join the list element into a single string
    string = ''.join(string_list_element)
    print(string)

    # Use regular expression to extract key-value pairs
    pairs = re.findall(r'(\w+):\s*\[([\d\s,]+)\]', string)
    print(pairs)

    # Initialize dictionary
    result_dict = {}

    # Iterate over pairs and populate dictionary
    for key, values_str in pairs:
        values = [int(value) for value in values_str.split(',') if value.strip()]
        result_dict[key] = values

    #print(result_dict)
        
    prog_bar_update_val += 1
    # records progress by updating prog bar with each file compiled
    window["-Progress_BAR-"].update(max = max, current_count=int(prog_bar_update_val))

    ###################### Use groups assignment dictionary to create an new df with optimal groups ######################

    # OG DataFrame with 'id' column
    for_optimal_groups_df = pd.read_excel(file_name)

    # Dictionary of group assignments
    group_assignments = result_dict

    # Function to map group assignments
    def map_group(id):
        for group, ids in group_assignments.items():
            if id in ids:
                return group
        return None

    # Apply mapping function to 'id' column to create 'group' column
    for_optimal_groups_df['group'] = for_optimal_groups_df['id'].apply(map_group)

    # Display the new DataFrame
    #print(for_optimal_groups_df)

    # Save the updated DataFrame to a new sheet in the same Excel file
    with pd.ExcelWriter(new_file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        for_optimal_groups_df.to_excel(writer, sheet_name='OptimalAssignment', index=False)

    # last prog bar addition indicating the end of the program run
    window["-Progress_BAR-"].update(current_count=int(prog_bar_update_val +1))

    ###################### Use 'OptimalGroup' df to graph these groups ######################
            
    df = for_optimal_groups_df

    # Extract time point columns
    time_point_columns = [col for col in df.columns if col not in ['group', 'combo', 'id', 'cage']]

    # Calculate average values for 'a' and 'b' groups for each column data point
    a_avg = df[df['group'] == 'a'][time_point_columns].mean()
    b_avg = df[df['group'] == 'b'][time_point_columns].mean()

    # Plot the data
    plt.figure(figsize=(10, 6))

    # Plot 'a' average data
    plt.plot(time_point_columns, 
            a_avg, marker='o', label='a', color='blue')

    # Plot 'b' average data
    plt.plot(time_point_columns, 
            b_avg, marker='o', label='b', color='red')

    plt.title('Group Avg\na vs b')
    plt.xlabel('Data Columns', labelpad=10)
    plt.ylabel('Behavior Units', labelpad=10)
    plt.legend()
    # Rotate x-axis labels by 90 degrees
    plt.xticks(rotation=90,  fontsize=8)
    plt.show()

    # window telling the user the program functioned correctly
    nom_window(new_file_name)

# main GUI creation and GUI elements
sg.theme('Reddit')

layout = [
    [sg.Text("Select the Excel file containing              \n"
             "behavioral data & subject id's                \n" 
             "to be assigned groups."),
    sg.Input(key="-IN-"),
    sg.FileBrowse()],

    [sg.Text("Select a file to store the new Excel doc. \n"
             "Original data will be copied & group      \n"
             "evaluation will be saved to this location.\n"),
    sg.Input(key="-OUT-"),
    sg.FolderBrowse()],

    [sg.Exit(), sg.Button("Press to assign subjects to least different groups"), 
    sg.Text("eBot's progress..."),
    sg.ProgressBar(20, orientation='horizontal', size=(15,10), 
                border_width=4, bar_color=("Blue", "Grey"),
                key="-Progress_BAR-")]
    
]

# create the window
window = sg.Window("Welcome to eBot's Least Diff Group Assignor!", layout)

# create an event loop
while True:
    event, values = window.read()
    # end program if user closes window
    if event == "Exit" or event == sg.WIN_CLOSED:
        break
    if event == "Press to assign subjects to least different groups":
        # check file selections are valid
        if (is_valid_path(values["-IN-"])) and (is_valid_path(values["-OUT-"])):

            extract_values_from_excel(
            input_filename  = values["-IN-"],
            output_folder = values["-OUT-"])   

window.close