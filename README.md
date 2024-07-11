# Email Analysis Tool

This Python script analyzes email data from an Excel file to compute the number of hours worked based on email activity within specified working hours. It outputs a summary of daily work hours, the timing of the first and last emails, and any extra hours worked outside of normal working hours.

## Prerequisites

Before running this script, ensure that you have Python and `pip` installed on your machine. This guide will use Homebrew for installation on macOS. If you are using another OS, you will need to adjust the installation commands accordingly.

### Install Python and pip

If you don't have Homebrew installed, open a terminal and run the following command:

```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

Once Homebrew is installed, you can install Python and Git:

```bash
brew install python git
```

This command installs the latest version of Python (which includes pip) and Git.

### Install Pandas via pip

You will need the Pandas library to run the script. It's recommended to manage Python packages via a virtual environment, but for simplicity, you can install Pandas globally:

```bash
pip install pandas
```

Alternatively, if the repository has a requirements.txt file specifying all dependencies, you can install all required packages with:

```bash
pip install -r requirements.txt
```

## Cloning the Repository

To get the script, you will need to clone the repository where it is hosted. Use the following Git command:

```bash
git clone https://github.com/AmauryGarde/email_work_hours_analysis.git
```

This command clones the repository to your local machine.

## Configuration

Before running the script, you will need to update it with specific details pertaining to your document:

1. Update the `document_name` variable with the path to your Excel file.
2. Ensure the `date_column` and `subject_column` match the column headers in your Excel file used for dates and email subjects respectively.

## Running the Script

To run the script, navigate to the folder containing your script and execute it with Python:

```bash
python email_analysis_tool.py
```

The script will read the specified Excel file, perform the analysis, and output the results to a new Excel file named `email_analysis.xlsx` in the same directory.

## Output

The output Excel file will contain the following columns:

- Date
- First email time
- First email subject
- Last email time
- Last email subject
- Total hours worked
- Extra hours worked

The final row of the Excel file provides a summary of total hours and extra hours worked for all dates combined.

