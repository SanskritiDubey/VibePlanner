import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import calendar


class LeavePlanner:
    def __init__(self, holidays_file, employees_file):
        """Initialize the leave planner with holiday and employee data."""
        try:
            self.holidays_df = pd.read_excel(holidays_file)
            print(f"Successfully loaded holiday data from {holidays_file}")
            print(f"Holiday columns: {self.holidays_df.columns.tolist()}")
        except Exception as e:
            print(f"Error loading holiday data: {e}")
            raise
            
        try:
            self.employees_df = pd.read_excel(employees_file)
            print(f"Successfully loaded employee data from {employees_file}")
            print(f"Employee columns: {self.employees_df.columns.tolist()}")
        except Exception as e:
            print(f"Error loading employee data: {e}")
            raise
            
        # Find city columns in the holiday dataframe
        self.cities = [col for col in self.holidays_df.columns if col not in ['SI No', 'Holiday Description', 'Date', 'Day']]
        print(f"Detected cities: {self.cities}")
        
        self.year = 2025  # Fixed to 2025 as per the calendar
        
        # Map employee dataframe columns to expected columns
        self.emp_name_col = self._find_column(['Employee Name', 'Name', 'EmployeeName', 'Employee', 'name'], self.employees_df)
        self.emp_id_col = self._find_column(['Employee ID', 'ID', 'EmployeeID', 'EmpID', 'employeeid'], self.employees_df)
        self.emp_city_col = self._find_column(['City', 'Location', 'Place', 'Office', 'state'], self.employees_df)
        self.emp_leaves_col = self._find_column(['Available Leaves', 'Leaves', 'LeaveBalance', 'Leave Balance', 'noofleaves'], self.employees_df)
        
        print(f"Using columns: Name={self.emp_name_col}, ID={self.emp_id_col}, City={self.emp_city_col}, Leaves={self.emp_leaves_col}")
        
    def _find_column(self, possible_names, df):
        """Find a column in the dataframe from a list of possible names."""
        for name in possible_names:
            if name in df.columns:
                return name
        # If no match found, return the first column as a fallback
        print(f"Warning: Could not find column matching {possible_names}. Using first column: {df.columns[0]}")
        return df.columns[0]
        
    def _is_weekend(self, date):
        """Check if a date is a weekend (Saturday or Sunday)."""
        return date.weekday() >= 5  # 5 = Saturday, 6 = Sunday
    
    def _get_holidays_for_city(self, city):
        """Get all holidays for a specific city."""
        # Filter rows where the city column has 'Holiday'
        city_holidays = self.holidays_df[self.holidays_df[city] == 'Holiday']
        # Convert date strings to datetime objects
        holidays = []
        for _, row in city_holidays.iterrows():
            date_str = row['Date']
            holiday_desc = row['Holiday Description']
            try:
                # Parse the date string to a datetime object
                date = pd.to_datetime(date_str, format='%d-%b')
                # Set the year to 2025
                date = date.replace(year=self.year)
                holidays.append({'date': date, 'description': holiday_desc})
            except Exception as e:
                print(f"Error parsing date {date_str}: {e}")
        return holidays
    
    def _find_optimal_leave_periods(self, city, available_leaves):
        """Find optimal periods to take leaves for a given city."""
        city_holidays = self._get_holidays_for_city(city)
        holiday_dates = [h['date'] for h in city_holidays]
        holiday_desc = {h['date']: h['description'] for h in city_holidays}
        
        # Create a calendar for the year with weekends and holidays marked
        calendar_days = {}
        start_date = datetime(self.year, 1, 1)
        end_date = datetime(self.year, 12, 31)
        current_date = start_date
        
        while current_date <= end_date:
            # Mark as holiday, weekend, or workday
            if current_date in holiday_dates:
                calendar_days[current_date] = {'type': 'Holiday', 'description': holiday_desc.get(current_date, 'Holiday')}
            elif self._is_weekend(current_date):
                calendar_days[current_date] = {'type': 'Weekend', 'description': 'Weekend'}
            else:
                calendar_days[current_date] = {'type': 'Workday', 'description': 'Workday'}
            current_date += timedelta(days=1)
        
        # Find clusters of holidays and weekends
        clusters = []
        current_cluster = []
        current_date = start_date
        
        while current_date <= end_date:
            if calendar_days[current_date]['type'] in ['Holiday', 'Weekend']:
                current_cluster.append(current_date)
            else:
                if len(current_cluster) > 0:
                    clusters.append(current_cluster)
                    current_cluster = []
            current_date += timedelta(days=1)
            
            # If we're at the end of the month, end the current cluster
            if current_date.day == 1 and len(current_cluster) > 0:
                clusters.append(current_cluster)
                current_cluster = []
        
        # Add the last cluster if it exists
        if len(current_cluster) > 0:
            clusters.append(current_cluster)
        
        # Find optimal leave periods by looking for workdays that bridge holidays/weekends
        leave_suggestions = []
        leaves_remaining = available_leaves
        
        # Sort clusters by potential value (longer clusters first)
        clusters.sort(key=len, reverse=True)
        
        for i, cluster in enumerate(clusters):
            if leaves_remaining <= 0:
                break
                
            # Look for workdays before and after the cluster that could be used as leaves
            cluster_start = min(cluster)
            cluster_end = max(cluster)
            
            # Check days before the cluster
            potential_before = []
            check_date = cluster_start - timedelta(days=1)
            while leaves_remaining > 0 and check_date >= start_date and calendar_days[check_date]['type'] == 'Workday':
                potential_before.append(check_date)
                check_date -= timedelta(days=1)
                if len(potential_before) >= 2:  # Limit to 2 days before
                    break
            
            # Check days after the cluster
            potential_after = []
            check_date = cluster_end + timedelta(days=1)
            while leaves_remaining > 0 and check_date <= end_date and calendar_days[check_date]['type'] == 'Workday':
                potential_after.append(check_date)
                check_date += timedelta(days=1)
                if len(potential_after) >= 2:  # Limit to 2 days after
                    break
            
            # Calculate the value of taking leaves (days off gained / leaves used)
            if potential_before or potential_after:
                leaves_to_use = min(leaves_remaining, len(potential_before) + len(potential_after))
                days_off = len(cluster) + leaves_to_use
                value = days_off / leaves_to_use if leaves_to_use > 0 else 0
                
                if value >= 1.5:  # Only suggest if there's good value
                    # Use leaves in order of best value
                    leaves_used = 0
                    leave_dates = []
                    
                    # Prioritize days that bridge gaps
                    all_potential = sorted(potential_before + potential_after)
                    for date in all_potential:
                        if leaves_used < leaves_to_use:
                            leave_dates.append(date)
                            leaves_used += 1
                    
                    if leave_dates:
                        # Get the holiday descriptions for the cluster
                        holiday_info = []
                        for date in cluster:
                            if calendar_days[date]['type'] == 'Holiday':
                                holiday_info.append(f"{date.strftime('%d-%b')}: {calendar_days[date]['description']}")
                        
                        total_period = sorted(cluster + leave_dates)
                        leave_suggestions.append({
                            'start_date': min(total_period),
                            'end_date': max(total_period),
                            'leaves_used': leaves_used,
                            'total_days_off': len(total_period),
                            'value': value,
                            'leave_dates': sorted(leave_dates),
                            'holiday_info': holiday_info
                        })
                        leaves_remaining -= leaves_used
        
        # Sort suggestions by value
        leave_suggestions.sort(key=lambda x: x['value'], reverse=True)
        return leave_suggestions
    
    def generate_leave_plans(self):
        """Generate leave plans for all employees."""
        results = []
        
        for _, employee in self.employees_df.iterrows():
            name = employee[self.emp_name_col]
            emp_id = employee[self.emp_id_col]
            city = employee[self.emp_city_col]
            
            # Handle potential non-numeric leave values
            try:
                available_leaves = int(employee[self.emp_leaves_col])
            except (ValueError, TypeError):
                print(f"Warning: Invalid leave value for {name}: {employee[self.emp_leaves_col]}. Using default of 10.")
                available_leaves = 10
            
            # Convert city to string to ensure it's not an integer
            city = str(city)
            
            # Check if city exists in holiday list
            if city not in self.cities:
                print(f"Warning: City '{city}' not found in holiday list. Available cities: {self.cities}")
                # Try to find a close match
                matched_city = None
                for holiday_city in self.cities:
                    # Convert holiday_city to string as well
                    holiday_city_str = str(holiday_city)
                    if city.lower() in holiday_city_str.lower() or holiday_city_str.lower() in city.lower():
                        matched_city = holiday_city
                        print(f"Found potential match: {city} -> {matched_city}")
                        break
                
                if matched_city:
                    city = matched_city
                else:
                    results.append({
                        'Employee Name': name,
                        'Employee ID': emp_id,
                        'City': city,
                        'Status': f"Error: City '{city}' not found in holiday list",
                        'Suggestions': []
                    })
                    continue
            
            leave_suggestions = self._find_optimal_leave_periods(city, available_leaves)
            
            results.append({
                'Employee Name': name,
                'Employee ID': emp_id,
                'City': city,
                'Available Leaves': available_leaves,
                'Suggestions': leave_suggestions
            })
        
        return results
    
    def save_suggestions_to_excel(self, output_file='leave_suggestions.xlsx'):
        """Save leave suggestions to an Excel file."""
        results = self.generate_leave_plans()
        
        # Create a DataFrame for the output
        output_rows = []
        
        for employee in results:
            name = employee['Employee Name']
            emp_id = employee['Employee ID']
            city = employee['City']
            available_leaves = employee.get('Available Leaves', 0)
            
            if not employee['Suggestions']:
                output_rows.append({
                    'Employee Name': name,
                    'Employee ID': emp_id,
                    'City': city,
                    'Available Leaves': available_leaves,
                    'Suggestion': 'No optimal leave periods found',
                    'Start Date': '',
                    'End Date': '',
                    'Leaves Required': '',
                    'Total Days Off': '',
                    'Leave Dates': '',
                    'Holiday Information': ''
                })
            else:
                # Add top 3 suggestions for each employee
                for i, suggestion in enumerate(employee['Suggestions'][:3]):
                    output_rows.append({
                        'Employee Name': name,
                        'Employee ID': emp_id,
                        'City': city,
                        'Available Leaves': available_leaves,
                        'Suggestion': f"Option {i+1}",
                        'Start Date': suggestion['start_date'].strftime('%d-%b-%Y'),
                        'End Date': suggestion['end_date'].strftime('%d-%b-%Y'),
                        'Leaves Required': suggestion['leaves_used'],
                        'Total Days Off': suggestion['total_days_off'],
                        'Leave Dates': ', '.join([d.strftime('%d-%b-%Y') for d in suggestion['leave_dates']]),
                        'Holiday Information': ', '.join(suggestion['holiday_info'])
                    })
        
        # Create and save the DataFrame
        output_df = pd.DataFrame(output_rows)
        output_df.to_excel(output_file, index=False)
        print(f"Leave suggestions saved to {output_file}")
        return output_df


if __name__ == '__main__':
    import sys
    import os
    
    # Default file names
    holidays_file = 'India+Holiday+Calendar+2025.xlsx'
    employees_file = 'employee_details.xlsx'
    
    # Check if files exist and allow command line arguments
    if len(sys.argv) > 1:
        holidays_file = sys.argv[1]
    if len(sys.argv) > 2:
        employees_file = sys.argv[2]
    
    # Verify files exist
    if not os.path.exists(holidays_file):
        print(f"Error: Holiday file '{holidays_file}' not found.")
        print(f"Current directory: {os.getcwd()}")
        print(f"Files in directory: {os.listdir()}")
        sys.exit(1)
        
    if not os.path.exists(employees_file):
        print(f"Error: Employee file '{employees_file}' not found.")
        print(f"Current directory: {os.getcwd()}")
        print(f"Files in directory: {os.listdir()}")
        sys.exit(1)
    
    print(f"\nProcessing holiday file: {holidays_file}")
    print(f"Processing employee file: {employees_file}\n")
    
    try:
        # Create the leave planner with the specific file names
        planner = LeavePlanner(holidays_file, employees_file)
        
        # Generate and save leave suggestions
        result_df = planner.save_suggestions_to_excel()
        
        print("\nLeave planning completed successfully!")
        print("\nTop suggestions for each employee:")
        
        # Display a summary of the results
        employees = result_df['Employee Name'].unique()
        for employee in employees:
            employee_suggestions = result_df[result_df['Employee Name'] == employee]
            print(f"\n{employee} ({employee_suggestions['City'].iloc[0]})")
            
            if 'No optimal leave periods found' in employee_suggestions['Suggestion'].values:
                print("  No optimal leave periods found")
            else:
                for _, row in employee_suggestions.iterrows():
                    print(f"  {row['Suggestion']}: {row['Start Date']} to {row['End Date']} ")
                    print(f"    Take leave on: {row['Leave Dates']}")
                    print(f"    Total days off: {row['Total Days Off']}")
                    print(f"    Holiday information: {row['Holiday Information']}")
    except Exception as e:
        print(f"\nError running leave planner: {e}")
        import traceback
        traceback.print_exc()
