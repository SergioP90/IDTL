# IDTL
**A CMD application to look up details about power distribution towers from the Imedexa catalogue**

![Version](https://img.shields.io/badge/Version-1.0-blue)
![Windows](https://img.shields.io/badge/Windows-COMPATIBLE-brightgreen) 
![Linux](https://img.shields.io/badge/Linux-UNTESTED-yellow) 
![Mac](https://img.shields.io/badge/MAC-INCOMPATIBLE-red) 

---

## Features
Legend:  
🟢 - Implemented  
🟡 - Implemented, has exceptions  
🔴 - Not implemented

| **Feature**                     | **Status** |
|---------------------------------|------------|
| Search basic details from ID    | 🟢        |
| Batch search from a list of IDs | 🟡        |
| Update catalogues               | 🟡        |
| Create excel files              | 🟢        |
| Extra calculations              | 🔴        |

---

## ⚠️ Exceptions on features
- **Batch search from a list of IDs**: Only supports `.txt` files with each ID placed on a different line, with the format AAA-b..b-c..c-Xdefg.
- **Update catalogues**: Catalogues need to be updated manually.

---

## 🛠 Requirements
- Python 3.12.4 or higher  
- cachetools 7.0.1
- colorama 0.4.6
- pandas 3.0.0

---

## 💾 Installation

### From source
1. Download [Python 3.12.4 or higher](https://www.python.org/) if you don't have it yet
2. Download the zip and extract (or clone using Git)  
3. Run `pip install -r requirements.txt` on a terminal 
    - *A virtual enviroment is highly recommended*
4. Set up the tables folder and settings file (You may download an example)
5. Run `main.py`  

### From exe
1. Download the zip and extract  
2. Set up the tables folder and settings file (You may download an example)
3. Run `IDTL.exe`  

---

## ⚙️ Configuring the settings file
The program uses a JSON settings file to locate the catalogue tables and define how data is read and exported.
The structure of this file **must match** your folder and table layout.

Example `settings.json`

```json
{
    "tables_folder": "tables",
    "tables": {
        "excavations": "excavation.xlsx",
        "shafts": "shaft.xlsx",
        "heads": "head.xlsx",
        "crosses": "cross.xlsx",
        "domes": "dome.xlsx"
    },
    "columns": {
        "id_name": "Type",
        "excavations": "Excavation",
        "shafts": "Shaft",
        "heads": "Head",
        "cross_a": "Cross_a",
        "cross_c": "Cross_c",
        "domes": "Dome"
    }
}
```

- `tables_folder`
    - Name of the main folder that contains all tower type subfolders.

- `tables`
    - Maps each table type to its corresponding .xlsx filename.
    - Filenames must match exactly (case-sensitive).

- `columns`
    - Defines the column names used when generating Excel output files.
    - These values can be customized freely.


Important notes:
- Do not add or remove fields from this file.
- If the program fails to read the settings file, regenerate it using `reset`.

---


## 📁 Folder structure
The program determines which tower types are available based on the folder structure inside the directory defined by `tables_folder`.

Each tower type must be represented by a folder named after its ID prefix.

For example, for the tower types:
- Condor -> CO
- Grand Condor -> GCO
- Icharous -> IC


Your structure must look like this (following the previous `settings.json` file example):

```bash
IDTL.exe (or main.py)
settings.json
tables/
  CO/
    cross.xlsx
    dome.xlsx
    excavation.xlsx
    head.xlsx
    shaft.xlsx
  GCO/
    cross.xlsx
    dome.xlsx
    excavation.xlsx
    head.xlsx
    shaft.xlsx
  IC/
    cross.xlsx
    dome.xlsx
    excavation.xlsx
    head.xlsx
    shaft.xlsx
```

Rules:
- Folder names must match the tower ID prefix (e.g. IC from IC-70000-20-N1444).
- Each folder must contain all .xlsx files defined in the tables section of the settings file.
- Filenames are case-sensitive.
- Extra files or folders will be ignored, but keeping the directory clean is recommended.

You can verify that your structure is valid by running `types`

The program will:
- Detect all folders with uppercase names as tower types,
- Check for the required table files.

---


## 📑 Table format
Each table file must be an .xlsx file containing a single sheet.
Only the first sheet is read by the program.

Formatting rules:
- The table must start at cell A1 (top-left corner), even if it is empty.
- Row 1 and Column A define the lookup axes.
- These axis cells must contain only numeric values.
- Axis labels (titles) must not be included.

Example of a valid table:

|       | 55000 | 70000 |
|-------|-------|-------|
| 15    | 10,26 | 12,67 |
| 20    | 10,57 | 13,21 |
| 25    | 10,9  | 13,8  |
| 30    | 11,27 | 14,28 |
| 35    | 11,82 | 14,96 |
| 40    | 12,1  | 15,32 |

Important:
- The orientation of the table does not matter. Tables may be rotated and will still be interpreted correctly (switching X and Y axes).
- Additional sheets are ignored.

---


## 💡 Example Usages

### Create or reset to default settings and view them
```bash
> reset
> settings
```

### View the tower types currently loaded in the program
```bash
> types
```

### Search a single ID
```bash
> lookup IC-70000-20-N1444
```

### Search a list of IDs "input.txt" and save them to an Excel "output.xlsx"
```bash
> file input.txt output.xlsx
```

### View more help about the program
```bash
> help
> help file
```