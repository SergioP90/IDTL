import os
import pandas as pd
from cachetools import LRUCache


class TableCache:
    def __init__(self, table_dir="tables", max_types_mem=5, table_names=None):
        if table_names is None:
            table_names = ["excavation.xlsx", "shaft.xlsx", "head.xlsx", "cross.xlsx", "dome.xlsx"]
            
        self.table_dir = table_dir
        self.cache = LRUCache(maxsize=max_types_mem)
        self.table_names = table_names


    def load_tables(self, type_name):
        """
        Load tables either from file, or cache if available
        
        Parameters:
            - type_name: The type to extract the tables from
        
        Returns:
            - List of pandas dataframes
        """
        tables = []

        if type_name in self.cache:
            tables = self.cache[type_name]
        else:
            type_path = os.path.join(self.table_dir, type_name)
            for filename in self.table_names:
                file_path = os.path.join(type_path, filename)
                df = pd.read_excel(file_path, index_col=0, header=0)
                df.index = df.index.astype(int)
                df.columns = df.columns.astype(int)
                tables.append(df)
        self.cache[type_name] = tables
        return tables