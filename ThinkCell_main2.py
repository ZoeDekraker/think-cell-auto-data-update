import json
import subprocess
import csv
from collections import defaultdict


def generate_json_for_thinkcell(data, template_path):
    # Convert data to a defaultdict for easier access.
    data = defaultdict(lambda: defaultdict(float), data)

    # Sort years and gather unique sheep types.
    years = sorted(data.keys())
    sheep_types = sorted(set(type for year in data for type in data[year]))

    # Set up the basic structure of the JSON object.
    chart_data = {
        "template": template_path,
        "data": [
            {
                "name": "Chart1",
                "table": [
                    # Header row with year labels.
                    [{"string": "Type"}] + \
                    [{"string": str(year)} for year in years],
                    # Empty row after the header.
                    [{"string": ""}] * (len(years) + 1)
                ] + [
                    # Rows for each sheep type with their data across all years.
                    [{"string": sheep_type}] + [{"number": data[year][sheep_type]}
                                                for year in years]
                    for sheep_type in sheep_types
                ]
            }
        ]
    }

    return json.dumps([chart_data], indent=4)


def run_thinkcell_cli(ppttc_file, output_pptx):
    command = [
        "C:\\Program Files (x86)\\think-cell\\ppttc.exe", ppttc_file, '-o', output_pptx]

    # Execute the command using subprocess.
    try:
        result = subprocess.run(
            command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        print("think-cell processing successful.")
    except subprocess.CalledProcessError as e:
        print("Error:", e)
        print("Standard Output:", e.stdout)
        print("Standard Error:", e.stderr)


def read_csv(data_file_path):
    data = {}
    with open(data_file_path, 'r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            year = row.pop('Year')
            data[year] = {sheep_type: float(count)
                          for sheep_type, count in row.items()}
    return data


template_path = "Herd_Growth.pptx"
ppttc_path = "Herd_Data_in_JSON.ppttc"
output_pptx_path = "Final_Sheep_Presentation.pptx"
data_file_path = "sheep.csv"

# upload data from csv
data = read_csv(data_file_path)

# Call the function and write the output to a file.
json_output = generate_json_for_thinkcell(data, template_path, ppttc_path)
with open(ppttc_path, 'w') as file:
    file.write(json_output)

# Run the think-cell CLI.
run_thinkcell_cli(ppttc_path, output_pptx_path)
