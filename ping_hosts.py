import pandas as pd
import subprocess
import os

# Input and output files
input_file = "Rancher_VMs.xlsx"  # Replace with your Excel file name
output_file = "ping_results.log"

# Read the Excel file
try:
    df = pd.read_excel(input_file, engine="openpyxl")
    hostnames = df["New Hostname"].dropna().tolist()
except Exception as e:
    print(f"Error reading the Excel file: {e}")
    exit(1)

# Open the log file
with open(output_file, "w") as log:
    for hostname in hostnames:
        try:
            result = subprocess.run(
                ["ping", "-c", "1", hostname],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            if result.returncode == 0:
                log.write(f"{hostname}: Success\n")
                print(f"{hostname}: Success")
            else:
                log.write(f"{hostname}: Failed\n")
                print(f"{hostname}: Failed")
        except Exception as e:
            log.write(f"{hostname}: Error ({e})\n")
            print(f"{hostname}: Error ({e})")

print(f"Ping results saved to {output_file}")
