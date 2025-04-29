import pandas as pd

def generate_timetable(input_file, entity_type, entity_value):
    # Read Excel file
    df = pd.read_excel(input_file)

    # Standardize
    df['Day'] = df['Day'].str.capitalize()
    df['Teacher'] = df['Teacher'].str.strip()
    df['Room'] = df['Room'].astype(str).str.strip()
    df['Time'] = df['Time'].astype(str).str.strip()

    # Choose filter column
    filter_column = 'Teacher' if entity_type == 'teacher' else 'Room'

    # Filter based on user input
    df_filtered = df[df[filter_column].str.lower() == entity_value.lower()]
    if df_filtered.empty:
        print(f"No data found for {entity_type}: {entity_value}")
        return

    # Pivot: Day as rows, Time as columns, show other info in cells
    cell_content = 'Room' if entity_type == 'teacher' else 'Teacher'
    timetable = df_filtered.pivot(index='Day', columns='Time', values=cell_content).fillna('')

    # Reorder days
    day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    timetable = timetable.reindex(day_order)

    # Optional: sort time columns chronologically
    try:
        time_order = sorted(timetable.columns, key=lambda x: pd.to_datetime(x.split('-')[0], format='%H:%M', errors='coerce'))
        timetable = timetable[time_order]
    except:
        pass  # fallback if time format fails

    # Reset for Excel output
    timetable = timetable.reset_index().rename(columns={'index': 'Day'})

    # Output file
    output_file = f"{entity_value.replace(' ', '_')}_timetable.xlsx"
    timetable.to_excel(output_file, index=False)
    print(f"Timetable saved as: {output_file}")

if __name__ == "__main__":
    input_file = "timetable.xlsx"
    choice = input("Generate timetable by 'teacher' or 'room'? ").strip().lower()
    
    if choice == 'teacher':
        value = input("Enter the teacher's name: ").strip()
    elif choice == 'room':
        value = input("Enter the room number: ").strip()
    else:
        print("Invalid option. Please type 'teacher' or 'room'.")
        exit()

    generate_timetable(input_file, choice, value)
