# Attendance Tracker Automation using Python

This is a real-world Python automation project that processes multiple Excel files, calculates attendance percentages, saves a clean summary report, logs every step, and runs automatically via Task Scheduler — all without manual intervention.

---

## Features

✅ Combines multiple Excel files from a folder  
✅ Cleans & organizes the data using `pandas`  
✅ Calculates employee-wise attendance percentage  
✅ Generates a final Excel summary report using `openpyxl`  
✅ Logs the automation process using `logging`  
✅ Scheduled to run weekly via Windows Task Scheduler

---

## Folder Structure

Attendance_Automation/
│
├── Data/ # Raw Excel attendance files
├── Reports/ # Final summary output
├── automation.log # Log file to track each run
├── attendance_tracker.py # Main script

---

## Tech Stack

- **Python 3.x**
- `pandas` – data manipulation  
- `openpyxl` – Excel file writing  
- `glob`, `os` – file handling  
- `logging` – status/error tracking  
- Windows Task Scheduler – scheduled execution

---

## How It Works

1. Drop raw attendance Excel files in the `/Data` folder
2. Run the `attendance_tracker.py` script (or let Task Scheduler handle it)
3. The script:
   - Reads and combines files
   - Calculates attendance %
   - Saves a summary report in `/Reports`
   - Writes logs into `automation.log`
4. The report is now ready — every week, automatically!

---

## Example Output

| Employee ID | Days Present | Total Days | Attendance % |
|-------------|--------------|-------------|---------------|
| EMP001      | 20           | 22          | 90.91%        |
| EMP002      | 17           | 22          | 77.27%        |


---

##Snapshots
<img width="1498" height="832" alt="Screenshot (11)" src="https://github.com/user-attachments/assets/2211be33-746e-43a8-a4b6-dd6c549e66d1" />
<img width="712" height="209" alt="Screenshot 2025-07-16 153018" src="https://github.com/user-attachments/assets/7c83abbb-0066-441c-b143-f80f1c90c169" />
<img width="923" height="203" alt="Screenshot 2025-07-16 153208" src="https://github.com/user-attachments/assets/43e875a0-a1dc-4b8d-be00-0f8c0bf5a68c" />


