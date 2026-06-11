def collect_data_sections(array_x, array_y, array_z):
    combined_array = array_x + array_y + array_z
    sorted_array = sorted(combined_array, key=lambda x: x[1])

    output = []
    current_section = None

    for data_point in sorted_array:
        value = data_point[0]
        start_time = data_point[1]
        end_time = data_point[2]

        if current_section is None:
            # Initialize a new section
            current_section = [[value, None, None], start_time, end_time]
        else:
            section_values = current_section[0]
            section_start_time = current_section[1]
            section_end_time = current_section[2]

            if start_time == section_end_time:
                # Continue the existing section
                section_values[1] = value
                current_section[2] = end_time  # Update the end time of the section
            else:
                # Start a new section
                section_values[2] = value
                output.append(current_section)
                current_section = [[value, None, None], start_time, end_time]

    # Add the last section to the output
    if current_section is not None:
        current_section[0][2] = current_section[0][0]  # Set the end value
        output.append(current_section)

    return output

array_x = [[1, 10, 15], [2, 20, 25], [3, 30, 35], [4, 40, 45]]
array_y = [[5, 12, 18], [6, 22, 26], [7, 33, 36]]
array_z = [[8, 10, 12], [9, 22, 28], [10, 32, 40]]

output = collect_data_sections(array_x, array_y, array_z)
print(output)