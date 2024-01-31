import pandas as pd
import numpy as np

def read_excel_file(file_path):
    df = pd.read_excel(file_path)
    return df

def assign_groups(students, group_size):
    num_students = len(students)
    num_groups = num_students // group_size
    remainder = num_students % group_size

    # Shuffle the list of students randomly
    np.random.shuffle(students)

    # Assign students to groups
    groups = [students[i * group_size: (i + 1) * group_size] for i in range(num_groups)]

    # Distribute remaining students to existing groups
    for i in range(remainder):
        groups[i % num_groups].append(students[num_groups * group_size + i])

    return groups

def write_to_excel(df, groups, output_file):
    # Create a new column in the DataFrame to store the assigned groups
    df['Assigned Group'] = np.nan

    # Update the DataFrame with the assigned groups
    for i, group in enumerate(groups):
        df.loc[df['Roll no.'].isin(group), 'Assigned Group'] = i + 1

    # Add a new 'Group' column to the DataFrame
    df['Group'] = df['Assigned Group']

    # Sort the DataFrame based on the 'Group' column
    df.sort_values(by='Group', inplace=True)

    # Write the updated DataFrame to a new Excel file
    df.to_excel(output_file, index=False)

def main():
    # Enter actual path to your Excel file
    file_path = input("Enter the path to your Excel file with .xlsx extension ")
    df = read_excel_file(file_path)

    # Extract the list of students from the DataFrame
    students = df['Roll no.'].tolist()

    # Get the desired group size from the user
    group_size = int(input("Enter the number of students per group: "))

    # Assign groups
    groups = assign_groups(students, group_size)

    # Write the assigned groups to a new Excel file
    output_file = 'output_groups.xlsx'
    write_to_excel(df, groups, output_file)

    print(f"Assigned groups written to {output_file}")

if __name__ == "__main__":
    main()
