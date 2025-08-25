"""
OTRS Ticket Data Analysis Script
Function: Analyze OTRS ticketing system Excel data and provide comprehensive ticket statistics and analysis reports

Main Features:
1. Read and analyze OTRS ticket data from Excel files
2. Automatically identify common OTRS column name variants
3. Provide multiple statistical analyses:
   - Count of records with empty FirstResponse (sorted by priority)
   - Priority distribution of currently open tickets
   - Age distribution analysis of open tickets
   - Daily new and closed ticket statistics
   - Ticket state distribution statistics
4. Automatically generate timestamped log files

Author: AI Assistant
Version: 1.3
Date: 2025-08-25
"""

import pandas as pd  # Data processing and analysis library
import numpy as np   # Numerical computing library
from datetime import datetime  # Date and time handling
import re  # Regular expression library for text matching and parsing
import sys  # System-related functions for standard output redirection
from collections import defaultdict  # Dictionary with default values

def analyze_otrs_tickets(file_path):
    """
    Main function for OTRS ticket data analysis
    Function: Read Excel file, identify column names, and perform basic statistical analysis
    
    Parameters:
    file_path: Excel file path
    
    Returns:
    Dictionary containing statistical results or original DataFrame (if column mapping fails)
    """
    
    # Read Excel file with error handling
    print("Reading Excel file...")
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading file: {e}")
        return
    
    # Display basic data information for debugging and verification
    # print(f"Total records: {len(df)}")
    # print(f"Columns: {list(df.columns)}")
    # print("\nFirst few rows:")
    # print(df.head())
    # print("\nData types:")
    # print(df.dtypes) 
    
    # Check for common OTRS column name variants
    # Define possible column mappings to support various naming conventions
    possible_columns = {
        'created': ['Created', 'CreateTime', 'Create Time', 'Date Created', 'created', 'creation_date'],
        'closed': ['Closed', 'CloseTime', 'Close Time', 'Date Closed', 'closed', 'close_date'],
        'state': ['State', 'Status', 'Ticket State', 'state', 'status'],
        'ticket_number': ['Ticket Number', 'TicketNumber', 'Number', 'ticket_number', 'id']
    }
    
    # Find actual column names using case-insensitive matching
    actual_columns = {}
    for key, possible_names in possible_columns.items():
        for col in df.columns:
            # Use case-insensitive matching to find column names
            if any(name.lower() in col.lower() for name in possible_names):
                actual_columns[key] = col
                break
    
    print(f"\nIdentified columns: {actual_columns}")
    
    # If no columns found, display all available columns for manual mapping
    if not actual_columns:
        print("\nColumn mapping needed. Available columns:")
        for i, col in enumerate(df.columns):
            print(f"{i}: {col}")
        return df
    
    # Execute ticket statistical analysis with identified columns
    return analyze_ticket_statistics(df, actual_columns)

def analyze_ticket_statistics(df, columns):
    """
    Perform ticket statistical analysis
    Function: Calculate daily new and closed ticket counts, and count current open tickets
    
    Parameters:
    df: pandas DataFrame containing ticket data
    columns: Dictionary containing identified column name mappings
    
    Returns:
    stats: Dictionary containing statistical results
    """
    stats = {}
    
    # STATISTICAL CALCULATION: Convert date columns to datetime format for time-based analysis
    # This enables proper date operations and filtering
    if 'created' in columns:
        df['created_date'] = pd.to_datetime(df[columns['created']], errors='coerce')
        
        # STATISTICAL CALCULATION: Count daily new tickets using value_counts()
        # This calculates the frequency distribution of tickets by creation date
        daily_new = df['created_date'].dt.date.value_counts().sort_index()
        stats['daily_new'] = daily_new
    
    if 'closed' in columns:
        # Handle closed date (may contain NaN values for open tickets)
        df['closed_date'] = pd.to_datetime(df[columns['closed']], errors='coerce')
        
        # STATISTICAL CALCULATION: Count daily closed tickets (excluding NaN values)
        # Filter out open tickets and calculate frequency distribution by close date
        closed_tickets = df[df['closed_date'].notna()]
        if not closed_tickets.empty:
            daily_closed = closed_tickets['closed_date'].dt.date.value_counts().sort_index()
            stats['daily_closed'] = daily_closed
    
    # Display daily new and closed tickets in same line format
    if 'daily_new' in stats and 'daily_closed' in stats:
        print("\nDaily Ticket Statistics:")
        print("Date\t\tNew\tClosed")
        print("-" * 30)
        
        # Get all unique dates from both series
        all_dates = sorted(set(stats['daily_new'].index) | set(stats['daily_closed'].index))
        
        for date in all_dates:
            new_count = stats['daily_new'].get(date, 0)
            closed_count = stats['daily_closed'].get(date, 0)
            print(f"{date}\t{new_count}\t{closed_count}")
    elif 'daily_new' in stats:
        print("\nDaily New Tickets:")
        for date, count in stats['daily_new'].items():
            print(f"{date}: {count}")
    elif 'daily_closed' in stats:
        print("\nDaily Closed Tickets:")
        for date, count in stats['daily_closed'].items():
            print(f"{date}: {count}")
    
    # STATISTICAL CALCULATION: Current open tickets analysis
    # Use closed field to identify open tickets (tickets with empty closed date are open)
    if 'closed' in columns:
        current_open = df[df['closed_date'].isna()]
        stats['current_open_count'] = len(current_open)
        print(f"\nCurrent Open Tickets: {len(current_open)}")
        
        # Also show state distribution for reference
        if 'state' in columns:
            state_counts = df[columns['state']].value_counts()
            print(f"\nTicket States Distribution:")
            for state, count in state_counts.items():
                print(f"{state}: {count}")
    else:
        print("Closed column not found - cannot determine open tickets")
    
    return stats

def parse_age_to_hours(age_str):
    """
    Parse Age string to total hours
    Function: Convert time strings in format "X h Y m" or "X d Y h" to total hours
    
    Parameters:
    age_str: Age string, e.g., "2 h 10 m" or "1 d 12 h"
    
    Returns:
    total_hours: Total hours (float)
    """
    # STATISTICAL CALCULATION: Handle NaN values by returning 0
    # This ensures mathematical operations don't fail on missing data
    if pd.isna(age_str):
        return 0
    
    # Convert to lowercase for consistent processing
    age_str = str(age_str).lower()
    days = 0
    hours = 0
    minutes = 0
    
    # STATISTICAL CALCULATION: Extract days using regex pattern matching
    # Pattern matches digits followed by 'd' (with optional whitespace)
    day_match = re.search(r'(\d+)\s*d', age_str)
    if day_match:
        days = int(day_match.group(1))
    
    # STATISTICAL CALCULATION: Extract hours using regex pattern matching
    # Pattern matches digits followed by 'h' (with optional whitespace)
    hour_match = re.search(r'(\d+)\s*h', age_str)
    if hour_match:
        hours = int(hour_match.group(1))
    
    # STATISTICAL CALCULATION: Extract minutes using regex pattern matching
    # Pattern matches digits followed by 'm' (with optional whitespace)
    minute_match = re.search(r'(\d+)\s*m', age_str)
    if minute_match:
        minutes = int(minute_match.group(1))
    
    # STATISTICAL CALCULATION: Convert all time components to total hours
    # Mathematical formula: (days * 24) + hours + (minutes / 60)
    # This provides a standardized metric for age comparison and analysis
    return (days * 24) + hours + (minutes / 60)


def analyze_firstresponse_empty(df):
    """
    Analyze records with empty FirstResponse
    Function: Count tickets with empty FirstResponse field, grouped by priority
    
    Parameters:
    df: pandas DataFrame containing ticket data
    """
    print("\n" + "=" * 50)
    print("FIRST RESPONSE EMPTY ANALYSIS")
    print("=" * 50)
    
    # Find the actual FirstResponse column name (case-insensitive)
    firstresponse_col = None
    for col in df.columns:
        if 'firstresponse' in col.lower():
            firstresponse_col = col
            break
    
    if firstresponse_col:
        print(f"Using column: {firstresponse_col}")
        
        # Check for both NaN values and empty strings
        nan_empty = df[firstresponse_col].isna()
        empty_strings = df[firstresponse_col] == ''
        
        # Combine both conditions for empty FirstResponse
        empty_firstresponse_all = df[nan_empty | empty_strings]
        
        # Filter out Closed and Resolved states
        if 'State' in df.columns:
            # Exclude tickets with state containing 'closed' or 'resolved'
            open_states = ~empty_firstresponse_all['State'].str.lower().str.contains('closed|resolved', na=False)
            empty_firstresponse = empty_firstresponse_all[open_states]
            
            print(f"Total records with empty FirstResponse (excluding Closed/Resolved): {len(empty_firstresponse)}")
            print(f" - Before filtering (all empty FirstResponse): {len(empty_firstresponse_all)}")
            print(f" - Excluded Closed/Resolved tickets: {len(empty_firstresponse_all) - len(empty_firstresponse)}")
        else:
            empty_firstresponse = empty_firstresponse_all
            print(f"Total records with empty FirstResponse: {len(empty_firstresponse)}")
            print("State column not available - cannot filter out Closed/Resolved tickets")
        
        print(f" - NaN values: {nan_empty.sum()}")
        print(f" - Empty strings: {empty_strings.sum()}")
        
        # Display detailed information for empty FirstResponse tickets
        print("\nDetailed Empty FirstResponse Tickets:")
        columns_to_show = ['Ticket Number', 'Age', 'Created', 'Closed', 'FirstLock', firstresponse_col, 'State', 'Priority']
        available_columns = [col for col in columns_to_show if col in df.columns]
        
        if available_columns:
            detailed_info = empty_firstresponse[available_columns]
            print(detailed_info.to_string(index=False))
        else:
            print("Required columns not available for detailed view")
        
        # STATISTICAL CALCULATION: Group by priority if available
        if 'Priority' in df.columns:
            # Calculate frequency distribution of empty FirstResponse by priority
            priority_counts = empty_firstresponse['Priority'].value_counts()
            print("\nEmpty FirstResponse by Priority:")
            
            # STATISTICAL CALCULATION: Sort by priority in 1/2/3 order using custom mapping
            # This ensures consistent ordering regardless of original data order
            priority_order = {'1 very high': 1, '2 high': 2, '3 normal': 3}
            sorted_priorities = sorted(priority_counts.items(), 
                                     key=lambda x: priority_order.get(x[0], 999))
            
            # Output sorted priority statistics
            for priority, count in sorted_priorities:
                print(f"{priority}: {count}")
        else:
            print("Priority column not available for analysis")
    else:
        print("FirstResponse column not found in data")
        print("Available columns:", list(df.columns))

def analyze_open_tickets_by_priority(df, closed_column):
    """
    Analyze open tickets by priority distribution
    Function: Count the number of each priority level among currently open tickets
    
    Parameters:
    df: pandas DataFrame containing ticket data
    closed_column: Column name used to determine if a ticket is closed
    
    Returns:
    open_tickets: DataFrame containing all open tickets
    """
    print("\n" + "=" * 50)
    print("OPEN TICKETS BY PRIORITY DISTRIBUTION")
    print("=" * 50)
    
    # STATISTICAL CALCULATION: Identify open tickets using the closed column
    # Tickets are considered open if the closed field is empty (NaN)
    open_tickets = df[df[closed_column].isna()]
    print(f"Total open tickets: {len(open_tickets)}")
    
    # STATISTICAL CALCULATION: Analyze priority distribution if Priority column exists
    if 'Priority' in df.columns:
        # Calculate frequency distribution of priorities among open tickets
        priority_counts = open_tickets['Priority'].value_counts()
        print("\nOpen tickets by Priority:")
        for priority, count in priority_counts.items():
            print(f"{priority}: {count}")
    else:
        print("Priority column not available for analysis")
    
    return open_tickets

def analyze_open_tickets_by_age(df, open_tickets):
    """
    Analyze open tickets by age distribution
    Function: Count open tickets distributed across different age ranges
    
    Parameters:
    df: pandas DataFrame containing ticket data
    open_tickets: DataFrame containing all open tickets
    """
    print("\n" + "=" * 50)
    print("OPEN TICKETS BY AGE DISTRIBUTION ANALYSIS")
    print("=" * 50)
    
    # STATISTICAL CALCULATION: Check if Age column and parsed age_hours column exist
    if 'Age' in df.columns and 'age_hours' in df.columns:
        # Filter for open tickets with valid age data (exclude NaN values)
        open_with_age = open_tickets[open_tickets['Age'].notna()]
        
        # STATISTICAL CALCULATION: Count tickets in different age ranges
        # Using boolean indexing to filter tickets by age thresholds
        less_24h = len(open_with_age[open_with_age['age_hours'] <= 24])
        between_24_48h = len(open_with_age[(open_with_age['age_hours'] > 24) & (open_with_age['age_hours'] <= 48)])
        between_48_72h = len(open_with_age[(open_with_age['age_hours'] > 48) & (open_with_age['age_hours'] <= 72)])
        over_72h = len(open_with_age[open_with_age['age_hours'] > 72])
        
        # Output age distribution statistics
        print(f"Open tickets < 24 hours: {less_24h}")
        print(f"Open tickets 24-48 hours: {between_24_48h}")
        print(f"Open tickets 48-72 hours: {between_48_72h}")
        print(f"Open tickets > 72 hours: {over_72h}")
        
        # STATISTICAL CALCULATION: Count new tickets created today
        today = datetime.now().date()
        if 'created_date' in df.columns:
            # Count ALL tickets created today, not just open ones
            daily_new = len(df[df['created_date'].dt.date == today])
            print(f"New tickets today: {daily_new}")
        else:
            print("Created date information not available for daily count")
        
        # STATISTICAL CALCULATION: Count closed tickets today
        if 'Closed' in df.columns:
            # Convert Closed column to datetime if not already done
            if df['Closed'].dtype == 'object':
                df['closed_date_temp'] = pd.to_datetime(df['Closed'], errors='coerce')
            else:
                df['closed_date_temp'] = df['Closed']
            
            # Count tickets closed today (exclude NaN values)
            closed_today = len(df[df['closed_date_temp'].notna() & (df['closed_date_temp'].dt.date == today)])
            print(f"Closed tickets today: {closed_today}")
        else:
            print("Closed date information not available for today's count")
    else:
        print("Age data not available for analysis")

def prepare_data(file_path):
    """
    Prepare and process OTRS ticket data for analysis
    Function: Read Excel file, identify columns, and perform data transformations
    
    Parameters:
    file_path: Excel file path
    
    Returns:
    tuple: (df, closed_column, result) - DataFrame, closed column name, and analysis results
    """
    try:
        # Read Excel file
        df = pd.read_excel(file_path)
        print(f"Reading Excel file: {file_path}")
        print(f"Total records: {len(df)}")
        
        # Find closed column for open ticket identification
        closed_column = None
        possible_closed_names = ['Closed', 'CloseTime', 'Close Time', 'Date Closed', 'closed', 'close_date']
        for col in df.columns:
            if any(name.lower() in col.lower() for name in possible_closed_names):
                closed_column = col
                break
        
        # Fallback strategy if closed column not found
        if closed_column is None:
            print("Warning: Closed column not found, using State column for open ticket identification")
            closed_column = 'State'  # Fallback: use state-based identification
        
        # Parse Age column to hours for age-based analysis
        if 'Age' in df.columns:
            df['age_hours'] = df['Age'].apply(parse_age_to_hours)
        
        # Convert date columns for time-based analysis
        if 'Created' in df.columns:
            df['created_date'] = pd.to_datetime(df['Created'], errors='coerce')
        
        # Get analysis summary from original function
        result = analyze_otrs_tickets(file_path)
        
        return df, closed_column, result
        
    except Exception as e:
        print(f"Error processing file: {e}")
        return None, None, None

def generate_output(df, closed_column, result):
    """
    Generate comprehensive analysis output
    Function: Execute all analysis functions and produce formatted output
    
    Parameters:
    df: pandas DataFrame containing processed ticket data
    closed_column: Column name used for open ticket identification
    result: Dictionary containing basic analysis results
    """
    if df is None or result is None:
        return
    
    # 1. Output analysis summary
    if isinstance(result, dict):
        print("\n" + "=" * 50)
        print("ANALYSIS SUMMARY")
        print("=" * 50)
        
        # Summary statistics calculations
        # Count tickets created today (regardless of state)
        today = datetime.now().date()
        if 'created_date' in df.columns:
            today_new = len(df[df['created_date'].dt.date == today])
            print(f"New tickets today: {today_new}")
        else:
            print("Created date information not available for today's count")
        
        if 'daily_closed' in result:
            total_closed = result['daily_closed'].sum()
            print(f"Total closed tickets: {total_closed}")
        
        if 'current_open_count' in result:
            print(f"Current open tickets: {result['current_open_count']}")
    
    # 2. FirstResponse empty records analysis
    analyze_firstresponse_empty(df)
    
    # 3. Current open tickets by priority count
    open_tickets = analyze_open_tickets_by_priority(df, closed_column)
    
    # 4. Current open tickets age distribution analysis
    analyze_open_tickets_by_age(df, open_tickets)

if __name__ == "__main__":
    """
    Main program entry point
    Function: Execute complete OTRS ticket analysis workflow
    
    Usage: python analyze_tickets.py [excel_file_path]
    If no file path is provided, uses default file: ticket_search_2025-08-25_00-08.xlsx
    """
    # Handle command line arguments for file path
    import argparse
    
    parser = argparse.ArgumentParser(description='Analyze OTRS ticket data from Excel files')
    parser.add_argument('file_path', nargs='?', default="ticket_search_2025-08-25_00-08.xlsx",
                       help='Path to the Excel file to analyze (optional, uses default if not provided)')
    
    args = parser.parse_args()
    file_path = args.file_path
    
    # Create timestamped log file (Windows doesn't allow colons in filenames)
    log_filename = datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + ".txt"
    
    # Redirect all output to log file for comprehensive reporting
    original_stdout = sys.stdout
    with open(log_filename, 'w', encoding='utf-8') as f:
        sys.stdout = f
        
        print("OTRS TICKET ANALYSIS")
        print("=" * 50)
        
        # PHASE 1: DATA PREPARATION
        df, closed_column, result = prepare_data(file_path)
        
        # PHASE 2: OUTPUT GENERATION
        if df is not None:
            generate_output(df, closed_column, result)
        
        # Restore standard output
        sys.stdout = original_stdout
    
    print(f"Analysis completed. Results saved to: {log_filename}")
