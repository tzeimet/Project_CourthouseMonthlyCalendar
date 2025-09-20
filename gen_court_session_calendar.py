# gen_court_session_calendar.py 20250920

import configparser
from enum import Enum
from icecream import ic
from loguru import logger
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string, coordinate_to_tuple
import os
from pathlib import Path
import sys
import typer
import yaml

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
        
        # Open a new workbook.
        wb = Workbook()
        # Open the active worksheet. This would be the first of the new workbook.
        ws = wb.active
        
        #breakpoint()
        for k,v in yaml_config['worksheet'].items():
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
                color=yaml_config['ARGBColors'][v['font']['color']]  # Font color (Hex code - FF0000 is Red)
            )
            # Define fill color
            fill_color = PatternFill(
                start_color=yaml_config['ARGBColors'][v['fill']['start_color']], 
                end_color=yaml_config['ARGBColors'][v['fill']['end_color']], 
                fill_type=v['fill']['fill_type']
            )
            # Define border
            border = Border(
                left=Side(
                    style=v['border']['left']['style'], 
                    color=yaml_config['ARGBColors'][v['border']['left']['color']]
                ), 
                right=Side(
                    style=v['border']['right']['style'], 
                    color=yaml_config['ARGBColors'][v['border']['right']['color']]
                ), 
                top=Side(
                    style=v['border']['top']['style'], 
                    color=yaml_config['ARGBColors'][v['border']['top']['color']]
                ), 
                bottom=Side(
                    style=v['border']['bottom']['style'], 
                    color=yaml_config['ARGBColors'][v['border']['bottom']['color']]
                ) 
            )
            # Merge cells.
            if merge_cells in ['ByColumn','ByBoth',]:
                ws.merge_cells(f"{top_left_cell}:{bottom_right_cell}")
                # Set text
                ws[top_left_cell] = v['text'][0]
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
                    ws.cell(row=min_row, column=col).value = v['text'][i]
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
        
        wb.save(f"{Path(__file__).stem}.xlsx")

    except KeyError as e:
        logger.exception(f"FATAL: Missing required configuration key: {e}")
    except configparser.NoSectionError as e:
        logger.exception(f"FATAL: Configuration is missing required section(s). Details: {e}")
    except KeyboardInterrupt:
        logger.info("Program interrupted by user.")
    except Exception as e:
        # Catch-all block: Logs the exception and terminates.
        logger.exception("FATAL: Application terminated due to an unexpected unhandled exception.")
#==================================================================================================
if __name__ == "__main__":
    typer.run(main)
