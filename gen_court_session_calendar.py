# gen_court_session_calendar.py 20250921

import calendar
import configparser
from datetime import date, timedelta
from enum import Enum
from icecream import ic
from loguru import logger
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell import Cell
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string, coordinate_to_tuple
import os
from pathlib import Path
import sys
import typer
import yaml

#==============================================================================
# Constants
#==============================================================================
MAX_ROW = None
#------------------------------------------------------------------------------
#==============================================================================
# Classes
#==============================================================================
class SourcePDFError(Exception):
    """The source PDF causes an error."""
    pass
#------------------------------------------------------------------------------
#==============================================================================
# Functions
#==============================================================================
def setup_logging(config: configparser.ConfigParser):
    logs_folder = config['Paths']['logs_folder']
    # console logging
    console_log_level = config['Logging']['console_level']
    # Define a custom format string with color tags for the console output
    console_format = "<light-green>{time:YYYY-MM-DD HH:mm:ss}</light-green> | <level>{level: <8}</level> | <bold><cyan>{function:<20}</cyan>:<cyan>{line:>5}</cyan></bold> - <level>{message}</level>"
    console_format = config['Logging'].get('console_format',console_format)
    # file logging
    file_log_level = config['Logging']['file_level']
    # Define a custom log file path with a dynamic date
    file_log_path = Path(logs_folder) / Path(f"{Path(__file__).stem}"+"_{time:YYYY-MM-DD}.log")
    # Define a simple format string for the log file (no colors needed)
    file_format = "{time:YYYY-MM-DD HH:mm:ss} | {level:<8} | {function:<20}:{line:>5} - {message}"
    file_format = config['Logging'].get('file_format',file_format)
    # Remove the default handler
    logger.remove(0)
    # Add a sink for the console with the colored format
    logger.add(
        sys.stderr, # Use sys.stderr for console output
        level=console_log_level.upper(), 
        format=console_format,
        backtrace=False,
        diagnose=False,
        enqueue=True
    )
    # Add a sink for the log file with the plain format, daily rotation, and retention
    logger.add(
        file_log_path, 
        level=file_log_level.upper(), 
        rotation="1 day", 
        retention="1 week",
        format=file_format,
        backtrace=True,
        diagnose=True,
        enqueue=True
    )
    return logger
# =================================================================================================
# Read applicaytion configuration from a YAML file.
# =================================================================================================
def read_yaml_configuration(yamlFilename: str):
    """
    Reads the applicaiton's configuration informaiton from the given filename, yamlFilename. 
    Returns: A tuple (
        dfFinCols: Dictionary containing the Finances DataFrame structure, 
        # This is definition of the DataFrame, with column names and types, that the input CSV data 
        # files will be converted to.
        sourcedata: An array containg a dictionary for each type of input data file
        # Each dictionary defines how to convert its corresponding CSV data file into the finances DataFrame.
    )
    """
    ymldata = None
    with open(yamlFilename, 'r') as file:
        ymldata = yaml.safe_load(file)
    return(ymldata)
#--------------------------------------------------------------------------------------------------
def find_first_matching_cell(sheet: Worksheet , column_letter: str, text: str, row_max: int = MAX_ROW):
    """
    Finds the first row index (1-based) with the given text in the given column.
    """
    row_num = 1  # Start checking from the first row
    while sheet[f'{column_letter}{row_num}'].value != text and row_num < row_max :
        row_num += 1
    if not row_num < row_max:
        return None
    # Return the cell object or the row number
    # return sheet[f'{column_letter}{row_num}']
    return row_num
#--------------------------------------------------------------------------------------------------
def get_weekday_number(year: int, month: int, day: int):
    """
    Returns the weekday as an integer (1=Monday, 5=Sunday).
    """
    try:
        # Create a date object
        d = date(year, month, day)
        
        # Use the .weekday() method
        return d.isoweekday()
    except ValueError as e:
        # Handles errors like an invalid date (e.g., February 30th)
        print(f"Error creating date: {e}")
        return None
#--------------------------------------------------------------------------------------------------
def copy_cell(source_cell: Cell, dest_cell: Cell):
    """
    Copies the value, number format, and style properties of one cell to another.
    """
    # --- 1. Copy Value and Type ---
    dest_cell.value = source_cell.value
    # --- 2. Copy Number Formatting ---
    # This handles currency, date formats, percentages, etc.
    dest_cell.number_format = source_cell.number_format
    # --- 3. Copy Basic Styles (Font, Fill, Border, Alignment, etc.) ---
    # Note: openpyxl stores styles internally. Accessing the '_style' attribute 
    # provides the most reliable way to copy the entire style object.
    if source_cell.has_style:
         dest_cell._style = source_cell._style
    # --- 4. Copy Cell Protection (Locked/Hidden) ---
#    if source_cell.protection:
#        dest_cell.protection = source_cell.protection
#    # --- 5. Copy Comment (if present) ---
    if source_cell.comment:
        dest_cell.comment = source_cell.comment
#--------------------------------------------------------------------------------------------------
#==================================================================================================
def main(
    config_file: str = typer.Argument(
        f"{Path(__file__).stem}.ini"
        ,help=(
            ""
        )
    )
    ,yaml_config_file: str = typer.Argument(
        f"{Path(__file__).stem}.yaml"
        ,help=(
            ""
        )
    )
):
    """
    """
    try:
        # App initialization
        config = configparser.ConfigParser()
        config.read(config_file)
        for k,v in config['Paths'].items():
            path = Path(v)
            if not path.is_dir():
                path.mkdir()
        
        logger = setup_logging(config)
        logger.info(f"Starting: {Path(__file__).stem}")
        
        logger.info(f"Reading YAML configurations from: {yaml_config_file}")
        yaml_config = read_yaml_configuration(yaml_config_file)
        
        #  Data from YAML config
        calendar_year = yaml_config['data']['calendar_year']
        MAX_ROW = yaml_config['constants']['worksheet']['MAX_ROW']
        
        # Open a new workbook.
        wb = Workbook()
        # Open the active worksheet. This would be the first of the new workbook.
        ws = wb.active
        # Rename the worksheet
        sheet_name = yaml_config['worksheet']['sheet_name']
        ws.title = sheet_name.replace(
            "${calendar_month_name}$"
            ,calendar.month_name[1]
        ).replace(
            "${calendar_year}$"
            ,str(calendar_year)
        )
        
        # For each subkey of ['worksheet'], create the sheet's layout.
        for k,v in yaml_config['worksheet'].items():
            if not isinstance(v,dict):
                continue
            logger.info(f"Creating: {k}")
            top_left_cell = v['cell_range']['top_left_cell']
            bottom_right_cell = v['cell_range']['bottom_right_cell']
            min_row, min_col = coordinate_to_tuple(top_left_cell)
            max_row, max_col = coordinate_to_tuple(bottom_right_cell)
            merge_cells = v['cell_range']['merge_cells']
            # Define font
            font = Font(
                name=v['font']['name'],             # Font type/family
                size=v['font']['size'],             # Font size (points)
                bold=v['font']['bold'],             # Optional: Make text bold
                italic=v['font']['italic'],         # Optional: Make text italic
                color=yaml_config['constants']['ARGBColors'][v['font']['color']]  # Font color (Hex code - FF0000 is Red)
            )
            # Define fill color
            fill_color = PatternFill(
                start_color=yaml_config['constants']['ARGBColors'][v['fill']['start_color']], 
                end_color=yaml_config['constants']['ARGBColors'][v['fill']['end_color']], 
                fill_type=v['fill']['fill_type']
            )
            # Define border
            border = Border(
                left=Side(
                    style=v['border']['left']['style'], 
                    color=yaml_config['constants']['ARGBColors'][v['border']['left']['color']]
                ), 
                right=Side(
                    style=v['border']['right']['style'], 
                    color=yaml_config['constants']['ARGBColors'][v['border']['right']['color']]
                ), 
                top=Side(
                    style=v['border']['top']['style'], 
                    color=yaml_config['constants']['ARGBColors'][v['border']['top']['color']]
                ), 
                bottom=Side(
                    style=v['border']['bottom']['style'], 
                    color=yaml_config['constants']['ARGBColors'][v['border']['bottom']['color']]
                ) 
            )
            # Merge cells.
            if merge_cells in ['ByColumn','ByBoth',]:
                ws.merge_cells(f"{top_left_cell}:{bottom_right_cell}")
                # Set text
                ws[top_left_cell] = v['text'][0] if isinstance(v['text'],list) else v['text']
                # Set font
                ws[top_left_cell].font = font
                # Set alignment.
                ws[top_left_cell].alignment = Alignment(horizontal=v['alignment']['horizontal'], vertical=v['alignment']['vertical'])
                # Set fill color.
                ws[top_left_cell].fill = fill_color
            else: # 'ByRow'
                # Merge cells.
                for i,col in enumerate(range(min_col,max_col+1)):
                    ws.merge_cells(
                        start_row=min_row
                        ,start_column=col
                        ,end_row=max_row
                        ,end_column=col
                    )
                    # Set text
                    ws.cell(row=min_row, column=col).value = v['text'][i]  if isinstance(v['text'],list) else v['text']
                    # Set font
                    ws.cell(row=min_row, column=col).font = font
                    # Set alignment.
                    ws.cell(row=min_row, column=col).alignment = Alignment(horizontal=v['alignment']['horizontal'], vertical=v['alignment']['vertical'])
                    # Set fill color.
                    ws.cell(row=min_row, column=col).fill = fill_color
            # Add border
            # You must apply the border to ALL cells in the merged range
            # since a single cell's border won't cover the entire merged area.
            for row in ws.iter_rows(min_row=min_row, min_col=min_col, max_row=max_row, max_col=max_col):
                for cell in row:
                    cell.border = border
            # Set column width
            if v.get('column_width_inches',None) is not None:
                # Set width of columns
                cell_width = yaml_config['excel']['cell_unit_per_inch'] * v['column_width_inches']
                _, col_idx_left = coordinate_to_tuple(top_left_cell)
                _, col_idx_right = coordinate_to_tuple(bottom_right_cell)
                for col_idx in range(col_idx_left,col_idx_right+1):
                    ws.column_dimensions[get_column_letter(col_idx)].width = cell_width
            
        # Copy worksheet for remaining months
        yaml_config_ws_title = yaml_config['worksheet']['title']
        title_top_left_cell = yaml_config_ws_title['cell_range']['top_left_cell']
        title = yaml_config_ws_title['text'][0]  if isinstance(yaml_config_ws_title['text'],list) else yaml_config_ws_title['text']
        yaml_config_ws_subtitle = yaml_config['worksheet']['subtitle']
        subtitle_top_left_cell = yaml_config_ws_subtitle['cell_range']['top_left_cell']
        subtitle = yaml_config_ws_subtitle['text'][0]  if isinstance(yaml_config_ws_subtitle['text'],list) else yaml_config_ws_subtitle['text']
        judges = yaml_config['data']['judges']
        judges_count = len(judges)
        if judges_count < 1:
            raise Exception("Judges are not specified in YAML configuration file.")
        for m in range(2,13):
            # Copy worksheet
            new_ws = wb.copy_worksheet(ws)
            # Rename the new worksheet
            new_ws.title = sheet_name.replace("${calendar_month_name}$",calendar.month_name[m]).replace("${calendar_year}$",str(calendar_year))
            # Change title in sheet
            new_ws[title_top_left_cell] = title.replace("${calendar_month_name}$",calendar.month_name[m]).replace("${calendar_year}$",str(calendar_year))
            # Change subtitle in sheet
            new_ws[subtitle_top_left_cell] = subtitle.replace("${judge}$",judges[m - judges_count*int((m-1)/judges_count) - 1])
            #
        # Change title in first (January) sheet
        ws[title_top_left_cell] = title.replace("${calendar_month_name}$",calendar.month_name[1]).replace("${calendar_year}$",str(calendar_year))
        # Change subtitle in first (January) sheet
        ws[subtitle_top_left_cell] = subtitle.replace("${judge}$",judges[0])
        
        # For each sheet (month) set up month days and placeholders for court sessions.
        # The month day cells are indicated by '${calendar_day}$ placeholder.
        month_day_placeholder = '${calendar_day}$'  ### put in yaml_config
        for month in range(1,13):
            # Select the worksheet by index.
            ws = wb.worksheets[month-1]
            month_day = 1 #
            month_num_days = calendar.monthrange(calendar_year, month)[1]
            while True:
                # Locate top left month day placeholder.
                row_num = find_first_matching_cell(ws,'A',month_day_placeholder,100)
                if row_num is None:
                    raise Exception("Unable to locate top left month day placeholder.")
                # Add another week of rows to the sheet.
                ws.insert_rows(row_num+2, amount=2)
                # Copy cells to new rows.
                for row in range(row_num,row_num+2):
                    for col in range(1,6):
                        copy_cell(ws.cell(row,col),ws.cell(row+2,col))
                # Get the next month_day that is a workday.
                while (workday := get_weekday_number(calendar_year,month,month_day)) >= 5: # >= Friday
                    month_day += 1
                #
                for cell_workday in range(1,6):
                    if cell_workday < workday or month_day > month_num_days:
                        cell_value = ''
                    else:
                        cell_value = month_day
                        month_day += 1
                    ws.cell(row_num,cell_workday).value = cell_value
                if month_day+2 > month_num_days: # account for Sat & Sun
                    # Cleanup, Remove unneeded row.
                    row_num += 2
                    ws.delete_rows(row_num, amount=2)
                    break
            #break # month loop
        
        #sys.exit()
             
        workbook_name = f"{Path(__file__).stem}.xlsx"
        logger.info(f"Saving workbook: {workbook_name}")
        wb.save(workbook_name)

    except KeyError as e:
        logger.exception(f"FATAL: Missing required configuration key: {e}")
    except configparser.NoSectionError as e:
        logger.exception(f"FATAL: Configuration is missing required section(s). Details: {e}")
    except KeyboardInterrupt:
        logger.info("Program interrupted by user.")
    except Exception as e:
        # Catch-all block: Logs the exception and terminates.
        logger.exception(f"FATAL: Application terminated due to an unexpected unhandled exception: {e}")
#==================================================================================================
if __name__ == "__main__":
    typer.run(main)
