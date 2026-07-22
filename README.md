# Streamlit User Data Collection App

A web application built with Streamlit that collects user input through a form and saves the data to spreadsheets (CSV and Excel).

## Features

- ✅ **User Input Form**: Collects user information including name, email, phone, age, gender, location, and comments
- 💾 **Auto-Save**: Data is automatically saved to CSV file
- 📊 **View Data**: Display all collected data in an interactive table
- 📥 **Download Options**: Export data as CSV or Excel files
- 🗑️ **Clear Data**: Option to clear all saved data
- ⏰ **Timestamp**: Each entry is automatically timestamped
- 📱 **Responsive Design**: Works on desktop and mobile devices

## Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

## Installation

1. Navigate to the project directory:
   ```bash
   cd c:\Users\KMN\Desktop\Anji
   ```

2. Create a virtual environment (optional but recommended):
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```

3. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Running the Application

1. Start the Streamlit app:
   ```bash
   streamlit run app.py
   ```

2. The app will open in your default web browser at `http://localhost:8501`

3. Fill out the form and click "Submit Data"

4. View, download, or clear the collected data

## How It Works

1. **Data Input**: Users fill in the form with their information
2. **Validation**: Required fields (Name and Email) are validated
3. **Storage**: Data is saved to `user_data.csv` in the project directory
4. **Display**: All saved data is displayed in a table below the form
5. **Export**: Users can download data as CSV or Excel files

## File Structure

```
Anji/
├── app.py                 # Main Streamlit application
├── requirements.txt       # Python dependencies
├── README.md             # This file
└── user_data.csv         # Data file (created after first submission)
```

## Data Saved

Each entry contains:
- Timestamp
- Name
- Email
- Phone
- Age
- Gender
- City
- Country
- Comments

## Notes

- The `user_data.csv` file is created automatically in the project directory
- All data is stored locally on your machine
- You can back up the CSV file to prevent data loss
- The Excel export requires `openpyxl` library (included in requirements)

## Troubleshooting

**App doesn't start:**
- Make sure all packages are installed: `pip install -r requirements.txt`
- Check Python version: `python --version`

**Can't find the data file:**
- The `user_data.csv` is saved in the same directory as `app.py`
- Check if you have write permissions in the directory

**Excel download not working:**
- Make sure `openpyxl` is installed: `pip install openpyxl`

## License

Free to use and modify

## Support

For issues or questions, check the Streamlit documentation: https://docs.streamlit.io/
