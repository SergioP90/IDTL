import os
import pandas as pd


def safe_lookup(df, row_value, col_value):
    """
    Try to fetch df.at[row_value, col_value].
    If fails, try the transposed axes.
    Raises KeyError if both fail.
    """
    try:
        return df.at[row_value, col_value]
    except KeyError:
        try:
            return df.T.at[row_value, col_value]
        except KeyError:
            raise KeyError(f"Value not found in table: row={row_value}, col={col_value}")


def find_values(ID, table_cache):
    """
    Find all the required values for a given ID in the format \"A-B-C-XDEFG where:
        - A represents the type
        - B repreesents the effort
        - C represents the height
        - D represents the head type
        - E represents the cross type a
        - F represents the cross type c
        - G represents the dome type
        - X is a letter that will be ignored

    Paramenters:
        - ID: The ID to look up
        - table_cache: Cache object to locate and retrieve tables

    Returns
        - list [excavation, shaft, head, cross_a, cross_c, dome]
    """
    try:
        elements = ID.strip().split("-")
        if len(elements) < 4:
            raise ValueError("ID is not correctly formed")

        type_name = elements[0]
        effort = int(elements[1])
        height = int(elements[2])

        assembly = elements[3]
        if len(assembly) < 5:
            raise ValueError(f"Assembly part too short: {assembly}")

        head_type = int(assembly[1])
        cross_type_a = int(assembly[2])
        cross_type_c = int(assembly[3])
        dome_type = int(assembly[4])

        tables = table_cache.load_tables(type_name)

        excavation = safe_lookup(tables[0], height, effort)
        shaft = safe_lookup(tables[1], height, effort)
        head = safe_lookup(tables[2], head_type, effort)
        cross_a = safe_lookup(tables[3], cross_type_a, effort)
        cross_c = safe_lookup(tables[3], cross_type_c, effort)
        dome = safe_lookup(tables[4], dome_type, effort)

        return [excavation, shaft, head, cross_a, cross_c, dome]

    except ValueError as ve:
        raise ValueError(f"ID parsing error for '{ID}': {ve}") from ve
    except KeyError as ke:
        raise KeyError(f"Lookup error for '{ID}': {ke}") from ke
    except FileNotFoundError as fe:
        raise FileNotFoundError(f"Table missing for type '{type_name}' in ID '{ID}'") from fe
