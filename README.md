# Mentimeter Export Data Converter

A local web tool that converts the Excel survey file from Mentimeter into a clean, analysis-ready format that can be stored in RedCap.

---

## Requirements

- **Python 3.8 or later** — check with `python3 --version`
- No other software needed; all dependencies are installed into an isolated virtual environment.

---

## Setup (one time only)

Open a terminal, navigate to this folder, and run the four commands below.

```bash
# 1. Go to the project folder (adjust the path if needed!)
cd ~/Desktop/"BHCI Capstone Data Conversion"

# 2. Create a virtual environment
python3 -m venv venv

# 3. Activate it
source venv/bin/activate          # macOS / Linux
# venv\Scripts\activate           # Windows (use this line instead)

# 4. Install dependencies
pip install flask openpyxl
```

---

## Running the Tool

Each time you want to use the tool:

```bash
# From the project folder with the venv active:
source venv/bin/activate          # macOS / Linux
# venv\Scripts\activate           # Windows

python3 app.py
```

You should see output like:

```
 * Running on http://127.0.0.1:5050
```

Open your browser and go to **http://localhost:5050**.

To stop the server, press `Ctrl + C` in the terminal.

---

## Using the Tool

1. **Upload** data export Excel file (drag-and-drop or click Browse).
2. **Enter the domain name** — capitalization and extra spaces are ignored. Code Book:

   | Domain Name | Code |
   |---|---|
   | Strong Minds and Strong Bodies | D1 |
   | Positive Self-Worth | D2 |
   | Caring Families and Relationships | D3 |
   | Safety | D4 |
   | Vibrant Communities | D5 |
   | Healthy Environments | D6 |
   | Racial Justice, Equity, and Inclusion | D7 |
   | Fun and Happiness | D8 |

3. Click **Convert & Download**. The file `Voters_converted.xlsx` will download automatically.
4. Click **← Convert Another File** to reset the form and convert a new file.

---