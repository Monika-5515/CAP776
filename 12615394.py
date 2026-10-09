# MY data , My story
# This program reads my daily activity data from an excel file and calculates activity indices and relationship analysis.
import openpyxl

# User-defined function to calculate the average
def calculate_average(values): 
    total = sum(values)
    count = len(values)
    return total / count

# Function to check data continuity
def check_date_continuity(data): 
    dates = [row[0] for row in data]
    for i in range(1, len(dates)):
        difference = dates[i] - dates[i-1]
        if difference.days !=1:
            print("Data continuity check: Failed")   
    print("Data continuity check: Passed")

# Function to validate activitty time values
def validate_data(data):
    missing = 0
    invalid = 0
    for row in data:
        for i in range(1,10):
            value = row[i]
            if value is None or value == "":
                missing +=1
            elif not isinstance(value, (int, float)) or value < 0:
                invalid += 1
    print("Missing values:", missing)
    print("Invalid values:", invalid)

#Main function
def main():
    # File handling; open the excel file   
    file_name ="12615394.xlsx"  
    try:
        workbook = openpyxl.load_workbook(file_name, data_only=True)
        sheet = workbook.active
        print("File opened successfully")
    except FileNotFoundError:
        print("File not found")

    # Read activity data from the excel sheet
    for row in sheet.iter_rows(min_row=6, values_only=True):
        if row[0] is None:
            break
        print(row)   

    # Stores the valid records in a list
    data = []
    for row in sheet.iter_rows(min_row=6, values_only=True):
        if row[0] is None:
            break
        data.append(row)    
    print(data)       

    # Check data continuity 
    check_date_continuity(data)

    # Validate the data
    validate_data(data)  

    # Store individual columns in list
    dates = []
    sleep = []
    fitness = []
    study = []
    coding = [] 
    classes = []
    free = []
    for row in data:
        dates.append(row[0])
        sleep.append(row[1])
        fitness.append(row[2])
        study.append(row[3])
        coding.append(row[4])
        classes.append(row[5])
        free.append(row[9])    

    # Calculate activity indices
    # Calculate tech productivity(tpi)
    tpi = calculate_average(coding)
    print("TPI =", tpi)

    # Calculate academic activity(aai)
    total_academic = [study[i] + classes[i] for i in range(len(data))]
    aai = calculate_average(total_academic)
    print("AAI =", aai)

    # Calculate physical activity(phai)
    phai = calculate_average(fitness)
    print("PHAI =", phai)

    # Calculate sleep & recovery(sri)
    sri = calculate_average(sleep)
    print("SRI =", sri)

    # Calculate activity balance(abi)
    for row in data:
        free.append(0 if row[9] == "" or row[9] is None else float(row[9]))
    abi = calculate_average(free)
    print("ABI =", abi)

    # Calculate time utilization(tui)
    total_tracked = [row[8] for row in data]
    tui = calculate_average(total_tracked)
    print("TUI =", tui)

    # Convert qualitative  responses into numerical values ( python dictionary)
    feeling_mapping = {
        "Excellent": 5,
        "Good":4,
        "Neutral": 3,
        "Low": 2,
        "Stressed": 1
    }
    satisfaction_mapping = {
        "Very Satisfied": 5,"very satisfied": 5,
        "Satisfied": 4,
        "Neutral": 3,
        "Unsatisfied": 2,
        "Very Unsatisfied": 1
    }
    energy_mapping = {
        "High": 3,
        "Medium": 2,
        "Low": 1
    }

    # Calculate experience index(ei)
    total_feeling = sum(feeling_mapping[row[10]] for row in data)
    total_satisfaction = sum(satisfaction_mapping[row[11]] for row in data)
    total_energy = sum(energy_mapping[row[12]] for row in data)
    valid_days = len(data)
    ei = (total_feeling + total_satisfaction + total_energy) / (3 * valid_days)
    print("EI=", ei)

    # Calculate data continuity(dci)
    valid_recorded_days = len(data)
    expected_days = 40
    dci = (valid_recorded_days / expected_days) * 100
    print("DCI=", dci,"%")

    # Calculate personal activity index(pai)
    pai = (0.15*tpi)+(0.20*aai)+(0.15*phai)+(0.20*sri)+(0.15*tui)+(0.10*ei)+(0.05*dci)
    print("PAI =", pai)

    # Relationship analysis
    energy_values = []
    for row in data:
        energy_values.append(energy_mapping[row[12]])
    print(energy_values)

    # Coding <-> energy
    coding_energy = sum(coding[i] * energy_values[i] for i in range(len(data)))
    print("Coding Energy =", coding_energy)

    # Sleep <-> energy
    sleep_energy = sum(sleep[i] * energy_values[i] for i in range(len(data)))
    print("Sleep Energy =", sleep_energy)

    # Study <-> satisfaction
    study_satisfaction = sum(study[i] * satisfaction_mapping[data[i][11]]
                             for i in range(len(data)))
    print("Study Satisfaction =", study_satisfaction)

    # Interpretation of personal findings
    print("\nFINDINGS")
    print("My average coding time was", round(tpi, 2), "minutes per day.")
    print("My average study and class time was", round(aai, 2), "minutes per day.")
    print("My data continuity was", round(dci, 2), "%.")
    print("My data shows how I spent more time on academic activities than coding.")
    print("My data continuity was ", round(dci, 2), "% shows how consistently I recorded my activities.")

    # Close the excel workbook
    workbook.close()

#Run the main function
if __name__ == "__main__":
   main()






