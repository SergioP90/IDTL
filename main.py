import os
import json
import time
import cmd
import re
import shutil
import pandas as pd
from colorama import init, Fore, Style

from src.table_cache import TableCache
from src.table_handler import find_values
from src.file_handler import id_extract, write_excel

SETTINGS_FILE = "settings.json"

DEFAULT_JSON = {
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

CONTEXT = {
    "tables_folder": None,
    "table_names": None,
    "column_names": None,
    "table_cache": None
}

ALIASES = {
    'exit': ["quit", "q", "e", "x", "EOF"],
    'help': ['h', '?'],
    'clear': ['cls', 'clc'],
    'showsettings': ['show', 'context', 'settings', 'ss', 'view', 'vs', 'status'],
    'resetsettings': ['rsettings', 'reset', 'resets', 'rs'],
    'seetypes': ['types', 'available', 'st'],
    'lookup': ['singlelookup', 'slookup', 'sl', 'search'],
    'filelookup': ['batchlookup', 'searchfile', 'file', 'filel', 'files', 'batchl', 'fl', 'bl']
}

init(autoreset=True)

# ==== CMD FUNCTIONS =====

#TODO:
# - LOOKUP: SEARCH A SINGLE ID
# - FILE: SEARCH AN ENTIRE FILE AND MAKE AN EXCEL FILE

class DistributionTowerLookupCLI(cmd.Cmd):
    intro = f"{Fore.YELLOW}Imedexa Distribution Tower Lookup. Type help or ? to list commands.\n{Style.RESET_ALL}"
    prompt = f"{Fore.CYAN}[IDTL] >>> {Style.RESET_ALL}"

    # == Overrrides ==
    def default(self, line):
        cmd, *args = line.split()
        cmd = cmd.lower()

        # Find which real command this alias maps to
        for real_cmd, aliases in ALIASES.items():
            if cmd == real_cmd or cmd in aliases:
                func = getattr(self, f"do_{real_cmd}", None)
                if func:
                    return func(" ".join(args))
        print(f"{Fore.RED}Unknown command: {cmd}{Style.RESET_ALL}")


    def do_help(self, arg):
        if arg:
            # Specific help for one command (with aliases shown)
            cmd_name = self.resolve_alias(arg)
            if cmd_name:
                func = getattr(self, "do_" + cmd_name, None)
                if func and func.__doc__:
                    print(f"{Fore.CYAN}\nHelp for '{Fore.GREEN}{cmd_name}{Fore.CYAN}':{Style.RESET_ALL}")
                    print(f"{Fore.CYAN}{func.__doc__}{Style.RESET_ALL}")
                    aliases = ALIASES.get(cmd_name, [])
                    if aliases:
                        print(f"{Fore.CYAN}Aliases:{Style.RESET_ALL}", ", ".join(aliases))
                else:
                    print(f"{Fore.RED}No help available for {arg}{Style.RESET_ALL}")
            else:
                print(f"{Fore.RED}Unknown command: {arg}{Style.RESET_ALL}")
        else:
            # General help
            print(f"""{Fore.CYAN}\n
Thank you for using IDTL!

Make sure you have your settings file correct and have setup the tables in the corresponding folders.
You can check this settings file with {Fore.GREEN}settings{Fore.CYAN} and reset them to their default values with {Fore.GREEN}reset{Fore.CYAN}.

To set up the tables, make sure they match the names in the \"tables\" field in settings. In the folder specified under the \"tables_folder\" field, 
make sure to have a folder per type of tower that mathes their ID. Each of these folders needs to have 5 .xlsx files named correctly, and the tables
must have their upper left corner as the first cell of sheet1, even if this one is empty.

For example, to process IDs for CONDOR distribution towers, like CO-12000-24-N3565, the folder structure by default should look like this:

tables/
  CO/   <- This folder must match the "type" in the ID.
    cross.xlsx
    dome.xlsx
    excavation.xlsx
    head.xlsx
    shaft.xlsx

To process more types of tower, add more folders with their respective tables, and the same names. You can check what types are detected
or what is missing with {Fore.GREEN}types{Fore.CYAN}.

The \"columns\" field is used to indicate the name of the columns in the output .xlsx.

In this help menu, the available commands are listed with their aliases in parentheses.
            """)

            print(f"{Fore.CYAN}\nList of available commands:{Style.RESET_ALL}")

            commands = self.get_commands_with_aliases()

            # Format into columns
            width = shutil.get_terminal_size((80, 20)).columns
            col_width = max(len(c) for c in commands) + 4
            cols = max(1, width // col_width)

            for i, entry in enumerate(commands, 1):
                print(entry.ljust(col_width), end="")
                if i % cols == 0:
                    print()
            if len(commands) % cols != 0:
                print()

            print(f"\n{Fore.CYAN}Type {Fore.GREEN}help <command>{Style.RESET_ALL} {Fore.CYAN}for specific help and descriptions.{Style.RESET_ALL}")
    

    # == Helpers ==
    def get_commands_with_aliases(self):
        """Return a list of formatted command+aliases strings."""
        result = []
        seen = set()

        for name in self.get_names():
            if name.startswith("do_"):
                cmd_name = name[3:]
                if cmd_name in seen:
                    continue
                seen.add(cmd_name)

                aliases = ALIASES.get(cmd_name, [cmd_name])
                alias_list = [cmd_name] + aliases
                result.append(
                    Fore.GREEN + cmd_name + Style.RESET_ALL +
                    f" {Fore.YELLOW}(" + ", ".join(alias_list) + f"){Style.RESET_ALL}"
                )
        return sorted(result)


    def resolve_alias(self, name):
        """Return canonical command name for a given alias."""
        if hasattr(self, "do_" + name):
            return name
        for cmd_name, aliases in ALIASES.items():
            if name in aliases or name == cmd_name:
                return cmd_name
        return None

    
    # == Commands ==
    def do_exit(self, line):
        """
        Exit the IDTL.
        
        Example: 
            exit
        """
        print(f"{Fore.RED}Exiting IDTL...{Style.RESET_ALL}")
        return True

    
    def do_clear(self, line):
        """
        Clear the console screen.
        
        Example: 
            clear
        """
        os.system('cls' if os.name == 'nt' else 'clear')
        print(self.intro)


    def do_showsettings(self, line):
        """
        Show the current settings in settings.json and on the context

        Example:
            settings
        """

        # Read settings file
        if not os.path.exists(SETTINGS_FILE):
            print(f"{Fore.RED}No settings file detected! Reset the settings to recreate it.{Style.RESET_ALL}")
        else:
            settings = None
            try:
                with open(SETTINGS_FILE, 'r') as f:
                    settings = json.load(f)
            except json.JSONDecodeError as e:
                print(f"{Fore.RED}Error parsing {SETTINGS_FILE}: {e}\nReset the settings to recreate it.{Style.RESET_ALL}")
            except Exception as e:
                print(f"{Fore.RED}Error reading {SETTINGS_FILE}: {e}\nReset the settings to recreate it.{Style.RESET_ALL}")
            
            if settings != None:
                print(f"{Fore.GREEN}Settings in {SETTINGS_FILE}:{Style.RESET_ALL}")
                print(f"{Fore.GREEN}- Tables Folder: {settings.get("tables_folder", f"{Fore.GREEN}unknown{Style.RESET_ALL}")}{Style.RESET_ALL}\n")
                table_names = settings.get("tables", {})
                if not table_names:
                    print(f"{Fore.RED}Error reading table names.{Style.RESET_ALL}")
                else:
                    print(f"{Fore.GREEN}- Tables Names:{Style.RESET_ALL}")
                    print(f"{Fore.GREEN}-- Excavations: {table_names.get("excavations", f"{Fore.RED}unknown{Fore.GREEN}")}{Style.RESET_ALL}")
                    print(f"{Fore.GREEN}-- Shafts: {table_names.get("shafts", f"{Fore.RED}unknown{Fore.GREEN}")}{Style.RESET_ALL}")
                    print(f"{Fore.GREEN}-- Heads: {table_names.get("heads", f"{Fore.RED}unknown{Fore.GREEN}")}{Style.RESET_ALL}")
                    print(f"{Fore.GREEN}-- Crosses: {table_names.get("crosses", f"{Fore.RED}unknown{Fore.GREEN}")}{Style.RESET_ALL}")
                    print(f"{Fore.GREEN}-- Domes: {table_names.get("domes", f"{Fore.RED}unknown{Fore.GREEN}")}{Style.RESET_ALL}")
                    print()

                column_names = settings.get("columns", {})
                if not column_names:
                    print(f"{Fore.RED}Error reading column names.{Style.RESET_ALL}")
                else:
                    print(f"{Fore.GREEN}- Columns Names:{Style.RESET_ALL}")
                    print(f"{Fore.GREEN}-- ID Name: {column_names.get("id_name", f"{Fore.RED}unknown{Fore.GREEN}")}{Style.RESET_ALL}")
                    print(f"{Fore.GREEN}-- Excavations: {column_names.get("excavations", f"{Fore.RED}unknown{Fore.GREEN}")}{Style.RESET_ALL}")
                    print(f"{Fore.GREEN}-- Shafts: {column_names.get("shafts", f"{Fore.RED}unknown{Fore.GREEN}")}{Style.RESET_ALL}")
                    print(f"{Fore.GREEN}-- Heads: {column_names.get("heads", f"{Fore.RED}unknown{Fore.GREEN}")}{Style.RESET_ALL}")
                    print(f"{Fore.GREEN}-- Cross_a: {column_names.get("cross_a", f"{Fore.RED}unknown{Fore.GREEN}")}{Style.RESET_ALL}")
                    print(f"{Fore.GREEN}-- Cross_c: {column_names.get("cross_c", f"{Fore.RED}unknown{Fore.GREEN}")}{Style.RESET_ALL}")
                    print(f"{Fore.GREEN}-- Domes: {column_names.get("domes", f"{Fore.RED}unknown{Fore.GREEN}")}{Style.RESET_ALL}")
                    print()

        # Read existing settings
        print(f"{Fore.GREEN} ============================================== \n")
        if not CONTEXT:
            print(f"{Fore.RED}Error reading settings! Reset the settings to recreate it.{Style.RESET_ALL}")
        else:
            print(f"{Fore.GREEN}Settings in memory:{Style.RESET_ALL}")
            print(f"{Fore.GREEN}- Tables Folder: {CONTEXT.get("tables_folder", f"{Fore.GREEN}unknown{Style.RESET_ALL}")}{Style.RESET_ALL}\n")
            table_names = CONTEXT.get("table_names", [])
            if not table_names or len(table_names) != 5:
                print(f"{Fore.RED}Error reading table names.{Style.RESET_ALL}")
            else:
                print(f"{Fore.GREEN}- Tables Names:{Style.RESET_ALL}")
                print(f"{Fore.GREEN}-- Excavations: {table_names[0]}{Style.RESET_ALL}")
                print(f"{Fore.GREEN}-- Shafts: {table_names[1]}{Style.RESET_ALL}")
                print(f"{Fore.GREEN}-- Heads: {table_names[2]}{Style.RESET_ALL}")
                print(f"{Fore.GREEN}-- Crosses: {table_names[3]}{Style.RESET_ALL}")
                print(f"{Fore.GREEN}-- Domes: {table_names[4]}{Style.RESET_ALL}")
                print()

            column_names = CONTEXT.get("column_names", [])
            if not column_names or len(column_names) != 7:
                print(f"{Fore.RED}Error reading column names.{Style.RESET_ALL}")
            else:
                print(f"{Fore.GREEN}- Columns Names:{Style.RESET_ALL}")
                print(f"{Fore.GREEN}-- ID Name: {column_names[0]}{Style.RESET_ALL}")
                print(f"{Fore.GREEN}-- Excavations: {column_names[1]}{Style.RESET_ALL}")
                print(f"{Fore.GREEN}-- Shafts: {column_names[2]}{Style.RESET_ALL}")
                print(f"{Fore.GREEN}-- Heads: {column_names[3]}{Style.RESET_ALL}")
                print(f"{Fore.GREEN}-- Cross_a: {column_names[4]}{Style.RESET_ALL}")
                print(f"{Fore.GREEN}-- Cross_c: {column_names[5]}{Style.RESET_ALL}")
                print(f"{Fore.GREEN}-- Domes: {column_names[6]}{Style.RESET_ALL}")
                print()


    def do_resetsettings(self, line):
        """
        Reset the settings file and settings in memory

        Example:
            reset
        """
        if os.path.exists(SETTINGS_FILE):
            try:
                os.remove(SETTINGS_FILE)
                print(f"{Fore.RED}File {SETTINGS_FILE} has been removed.")
            except FileNotFoundError:
                print(f"{Fore.RED}File {SETTINGS_FILE} not found.")
            except Exception as e:
                print(f"{Fore.RED}An error occurred: {e}")

        setup_context()


    def do_seetypes(self, line):
        """
        View what types are available and if there are issues with the tables

        Example:
            types
        """

        tables_folder = CONTEXT.get("tables_folder")
        table_names = CONTEXT.get("table_names", [])
        types_list = []
        elements = []
        
        if not os.path.exists(tables_folder):
            print(f"{Fore.RED}Cannot find tables folder, make sure it exists and matches the name in the settings.json file.")
            return

        for name in os.listdir(tables_folder):
            full_path = os.path.join(tables_folder, name)
            if os.path.isdir(full_path) and re.fullmatch(r"[A-Z]+", name):
                types_list.append(name)
            else:
                elements.append(name)

        print(f"{Fore.GREEN}Found types in {tables_folder}:{Style.RESET_ALL}")
        for type_name in types_list:
            dir_path = os.path.join(tables_folder, type_name)
            tables = []
            remainder = []

            seen_tables = set()

            for name in os.listdir(dir_path):
                full_path = os.path.join(dir_path, name)

                if name in table_names:
                    if name not in seen_tables:
                        tables.append(name)
                        seen_tables.add(name)
                    else:
                        remainder.append(f"{name} (duplicate)")
                else:
                    remainder.append(name)

            if len(tables) >= 5:
                print(f"{Fore.GREEN}- {type_name} ({len(tables)} / {len(table_names)}){Style.RESET_ALL}")
            else:
                print(f"{Fore.GREEN}- {type_name} ({Fore.RED}{len(tables)}{Fore.GREEN} / {len(table_names)}) {Fore.RED}(Missing tables){Style.RESET_ALL}")
            for name in tables:
                print(f"{Fore.GREEN}-- {name}{Style.RESET_ALL}")
            for name in remainder:
                print(f"{Fore.RED}-- {name}{Style.RESET_ALL}")

        for name in elements:
            print(f"{Fore.RED}- {name} (Not a valid type){Style.RESET_ALL}")


    def do_lookup(self, line):
        """
        Retrieve the data from a single ID

        Example:
            lookup CO-15000-33-N3565
        """
        parts = line.split()
        if len(parts) == 1:
            input_id = parts[0]
        else:
            print(f"{Fore.RED}Usage: lookup <tower_id>{Style.RESET_ALL}")
            return

        try:
            values = find_values(input_id, CONTEXT.get("table_cache"))
            names = CONTEXT.get("column_names")
            print(f"{Fore.GREEN}Values for {input_id}:{Style.RESET_ALL}")
            for i in range(len(values)):
                print(f"{Fore.GREEN}- {names[i + 1]}: {values[i]}{Style.RESET_ALL}")
        except Exception as e:
            readable_err = readable_error(e)
            print(f"{Fore.RED}\"{input_id}\" had an issue: {readable_err}")

        
    def do_filelookup(self, line):
        """
        Retrieve the data from a list of IDs stored in a txt file, each being in one line

        Example:
            file input.txt output.xlsx
        """
        parts = line.split()
        if len(parts) == 2:
            input_file = parts[0]
            output_file = parts[1]
            if not input_file.endswith('.txt'):
                input_file += '.txt'
            if not output_file.endswith('.xlsx'):
                output_file += '.xlsx'
        else:
            print(f"{Fore.RED}Usage: filelookup <path_to_txt_file> <path_to_xlsx_file>{Style.RESET_ALL}")
            return

        ids = id_extract(input_file)
        if not ids:
            print(f"{Fore.RED}Error reading the input file, make sure the name is correct and it contains all IDs, each placed on a different line.{Style.RESET_ALL}")
            return

        df = batch_finder(ids, CONTEXT.get("table_cache"), CONTEXT.get("column_names"))

        try:
            write_excel(df, output_file)
        except FileNotFoundError as e:
            print(f"{Fore.RED}Error, cannot create or modify the file \"{output_file}\". Make sure it's not in use by another program and that the folder has write permissions.{Style.RESET_ALL}")
        except PermissionError as e:
            print(f"{Fore.RED}Error, program lacks the necessary permissions to write to \"{output_file}\".{Style.RESET_ALL}")
        except Exception as e:
            print(f"{Fore.RED}An error occurred: {e}")
        print()
# ========================


# ==== HELPER FUNCTIONS ====
def setup_context():
    """
    Set up the global CONTEXT dictionary by reading from settings.json.
    If settings.json doesn't exist, create it with default values.
    """
    global CONTEXT
    
    
    if not os.path.exists(SETTINGS_FILE):
        print(f"{Fore.YELLOW}Creating default settings file: {SETTINGS_FILE}")
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(DEFAULT_JSON, f, indent=4)
        settings = DEFAULT_JSON
    else:
        try:
            with open(SETTINGS_FILE, 'r') as f:
                settings = json.load(f)
        except json.JSONDecodeError as e:
            print(f"{Fore.RED}Error parsing {SETTINGS_FILE}: {e}\nUsing default settings.")
            settings = DEFAULT_JSON
        except Exception as e:
            print(f"{Fore.RED}Error reading {SETTINGS_FILE}: {e}\nUsing default settings.")
            settings = DEFAULT_JSON
    
    CONTEXT["tables_folder"] = settings.get("tables_folder", DEFAULT_JSON["tables_folder"])

    table_mapping = settings.get("tables", DEFAULT_JSON["tables"])

    ordered_table_keys = ["excavations", "shafts", "heads", "crosses", "domes"]
    CONTEXT["table_names"] = [
        table_mapping.get(key, DEFAULT_JSON["tables"][key]) 
        for key in ordered_table_keys
    ]
    
    columns = settings.get("columns", DEFAULT_JSON["columns"])
    CONTEXT["column_names"] = [
        columns.get("id_name", DEFAULT_JSON["columns"]["id_name"]),
        columns.get("excavations", DEFAULT_JSON["columns"]["excavations"]),
        columns.get("shafts", DEFAULT_JSON["columns"]["shafts"]),
        columns.get("heads", DEFAULT_JSON["columns"]["heads"]),
        columns.get("cross_a", DEFAULT_JSON["columns"]["cross_a"]),
        columns.get("cross_c", DEFAULT_JSON["columns"]["cross_c"]),
        columns.get("domes", DEFAULT_JSON["columns"]["domes"])
    ]

    try:
        CONTEXT["table_cache"] = TableCache(CONTEXT["tables_folder"], 5, CONTEXT["table_names"])
    except Exception as e:
        print(f"{Fore.RED}Error initializing table cache: {e}")
        CONTEXT["table_cache"] = None


def readable_error(msg):
    msg_lower = str(msg).lower()
    new_msg = "An unexpected error occurred. Please check the ID, folder structure and the tables."
    
    if "parsing error" in msg_lower or "invalid literal" in msg_lower:
        new_msg = "The ID format is incorrect. Please use TYPE-EFFORT-HEIGHT-XDEFG (e.g., CO-15000-33-N3565)."
    if "not found (type" in msg_lower or "no such file" in msg_lower:
        new_msg = "A required table is missing. Check that tables/<TYPE>/ contains all five .xlsx files named as in settings.json."
    if "could not be read as an excel table" in msg_lower:
        new_msg = "A table file could not be read. Make sure it is a valid .xlsx and isn't corrupted or open in another program."
    if "lookup error" in msg_lower or "not found in table" in msg_lower:
        new_msg = "A value could not be found in one of the tables. The ID may reference a value this type doesn't define, or the table may use a different label convention (meters vs ID numbers)."
    if "table missing" in msg_lower:
        new_msg = "The required tables for this type are missing. Please check that the folder and Excel files exist."
    return new_msg + "\n\tDetails: " + str(msg)


def batch_finder(ids, table_cache, column_names=["Type", "Excavation", "Shaft", "Head", "Cross_a", "Cross_c", "Dome"]):
    results = []

    for ID in ids:
        try:
            values = find_values(ID, table_cache)
            results.append([ID] + values)
            print(f"{Fore.GREEN}\"{ID}\" processed correctly.")
        except Exception as e:
            readable_err = readable_error(e)
            results.append([ID] + [None]*6)
            print(f"{Fore.RED}\"{ID}\" had an issue: {readable_err}")

    df = pd.DataFrame(results, columns=column_names)
    return df
# ==========================
        
    
if __name__ == "__main__":
    setup_context()
    DistributionTowerLookupCLI().cmdloop()
    time.sleep(0.1)