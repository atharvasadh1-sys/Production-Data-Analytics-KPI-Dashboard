#!/usr/bin/env python3
"""
KPI Engine for Machines Input/Output Data
Processes the cleaned data from script1 and calculates all required KPIs.

Input: processed_data_script1_final.xlsx
Output: kpi_output_script2_final.xlsx

Author: Senior Python Data Engineer
Date: 2026-08-04
Version: 2.0.0 - Pandas 3.x Compatible
"""

import pandas as pd
import numpy as np
import logging
from datetime import datetime, timedelta, time
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils.dataframe import dataframe_to_rows
from typing import Dict, List, Tuple, Optional, Any, Union
import warnings
warnings.filterwarnings('ignore')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('kpi_engine.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class KPIEngineConfig:
    """
    Configuration class for KPI calculations.
    """
    
    def __init__(self):
        # Shift definitions with break times
        self.shifts = {
            'Shift 1': {
                'start': time(6, 0, 0),
                'end': time(14, 0, 0),
                'break_start': time(10, 30, 0),
                'break_end': time(11, 0, 0),
                'display': '06:00-14:00',
                'break_display': '10:30-11:00'
            },
            'Shift 2': {
                'start': time(14, 0, 0),
                'end': time(22, 0, 0),
                'break_start': time(18, 30, 0),
                'break_end': time(19, 0, 0),
                'display': '14:00-22:00',
                'break_display': '18:30-19:00'
            },
            'Shift 3': {
                'start': time(22, 0, 0),
                'end': time(6, 0, 0),
                'break_start': time(2, 30, 0),
                'break_end': time(3, 0, 0),
                'display': '22:00-06:00',
                'break_display': '02:30-03:00'
            }
        }
        
        # Required operators per side
        self.required_input_operators = 3
        self.required_output_operators = 3
        
        # Thresholds
        self.min_event_duration_seconds = 30


class KPIEngine:
    """
    Main KPI calculation engine.
    """
    
    def __init__(self, config: KPIEngineConfig):
        self.config = config
        self.processed_df = None
        self.kpi_summary = None
        self.exception_report = None
        self.management_summary = None
        self.calculation_audit = []
        
        # KPI storage
        self.machine_kpis = {}
        self.operator_kpis = {}
        self.material_kpis = {}
        self.shift_kpis = {}
        
    def load_processed_data(self, input_file: str) -> pd.DataFrame:
        """
        Load processed data from script1 output.
        
        Args:
            input_file: Path to processed data file
            
        Returns:
            DataFrame with processed data
        """
        logger.info(f"Loading processed data from: {input_file}")
        
        try:
            # Load the processed data sheet
            self.processed_df = pd.read_excel(input_file, sheet_name='Processed Data')
            
            # Parse timestamps
            self.processed_df['Timestamp'] = pd.to_datetime(self.processed_df['Timestamp'])
            
            # Ensure numeric columns are correct type
            numeric_cols = [
                'Machine_ID', 'Person_Count', 'Person_Count_Imputed',
                'Duration_Seconds', 'Duration_Minutes', 
                'Break_Overlap_Seconds', 'Adjusted_Duration_Seconds',
                'Adjusted_Duration_Minutes'
            ]
            for col in numeric_cols:
                if col in self.processed_df.columns:
                    self.processed_df[col] = pd.to_numeric(self.processed_df[col], errors='coerce').fillna(0)
            
            # Convert boolean columns
            bool_cols = [
                'Has_Box', 'Has_Box_Imputed', 'Has_No_Box', 'Has_No_Box_Imputed',
                'Is_Feeding', 'Is_Bringing_Material', 'Operator_Present',
                'Operator_Not_Present', 'Material_Present_Operator_Absent',
                'Operator_Present_Material_Absent', 'Is_Imputed',
                'Is_Valid', 'Is_Break_Time'
            ]
            for col in bool_cols:
                if col in self.processed_df.columns:
                    self.processed_df[col] = self.processed_df[col].astype(bool)
            
            # Ensure Shift is categorical
            if 'Shift' in self.processed_df.columns:
                self.processed_df['Shift'] = pd.Categorical(
                    self.processed_df['Shift'],
                    categories=['Shift 1', 'Shift 2', 'Shift 3'],
                    ordered=True
                )
            
            logger.info(f"Loaded {len(self.processed_df)} processed records")
            return self.processed_df
            
        except Exception as e:
            logger.error(f"Error loading processed data: {str(e)}")
            raise
    
    def calculate_machine_kpis(self) -> Dict[str, Any]:
        """
        Calculate machine KPIs: Running time, Downtime, Utilization %, Availability %.
        
        Returns:
            Dictionary with machine KPI results
        """
        logger.info("Calculating machine KPIs...")
        
        results = {}
        
        # Group by Machine_ID and Shift
        for machine_id in self.processed_df['Machine_ID'].unique():
            machine_mask = self.processed_df['Machine_ID'] == machine_id
            machine_data = self.processed_df[machine_mask].copy()
            
            results[machine_id] = {}
            
            for shift in ['Shift 1', 'Shift 2', 'Shift 3']:
                shift_mask = machine_data['Shift'] == shift
                shift_data = machine_data[shift_mask].copy()
                
                if len(shift_data) == 0:
                    results[machine_id][shift] = {
                        'Running_Time_Hours': 0,
                        'Downtime_Hours': 0,
                        'Utilization_Percent': 0,
                        'Availability_Percent': 0,
                        'Total_Time_Hours': 0
                    }
                    continue
                
                # Calculate total adjusted duration for this shift
                total_duration = shift_data['Adjusted_Duration_Seconds'].sum()
                total_hours = total_duration / 3600
                
                # Calculate running time (when machine is working)
                running_mask = shift_data['Machine_Status_Imputed'] == 'Working'
                running_time = shift_data.loc[running_mask, 'Adjusted_Duration_Seconds'].sum()
                running_hours = running_time / 3600
                
                # Calculate downtime (when machine is not working)
                downtime_mask = shift_data['Machine_Status_Imputed'] == 'Not Working'
                downtime_time = shift_data.loc[downtime_mask, 'Adjusted_Duration_Seconds'].sum()
                downtime_hours = downtime_time / 3600
                
                # Calculate utilization %
                utilization = (running_hours / total_hours * 100) if total_hours > 0 else 0
                
                # Calculate availability %
                availability = (running_hours / (running_hours + downtime_hours) * 100) if (running_hours + downtime_hours) > 0 else 0
                
                results[machine_id][shift] = {
                    'Running_Time_Hours': round(running_hours, 2),
                    'Downtime_Hours': round(downtime_hours, 2),
                    'Utilization_Percent': round(utilization, 2),
                    'Availability_Percent': round(availability, 2),
                    'Total_Time_Hours': round(total_hours, 2)
                }
        
        self.machine_kpis = results
        logger.info("Machine KPIs calculated successfully")
        return results
    
    def calculate_operator_kpis(self) -> Dict[str, Any]:
        """
        Calculate operator KPIs: Input availability, Output availability, Shortage duration.
        
        Returns:
            Dictionary with operator KPI results
        """
        logger.info("Calculating operator KPIs...")
        
        results = {}
        
        for shift in ['Shift 1', 'Shift 2', 'Shift 3']:
            shift_data = self.processed_df[self.processed_df['Shift'] == shift].copy()
            
            if len(shift_data) == 0:
                results[shift] = {
                    'Input_Availability_Hours': 0,
                    'Output_Availability_Hours': 0,
                    'Input_Shortage_Hours': 0,
                    'Output_Shortage_Hours': 0,
                    'Input_Availability_Percent': 0,
                    'Output_Availability_Percent': 0
                }
                continue
            
            results[shift] = {}
            
            # Calculate input side operator availability
            input_mask = shift_data['Side'] == 'Input'
            input_data = shift_data[input_mask].copy()
            
            if len(input_data) > 0:
                # Total time for input side
                total_input_time = input_data['Adjusted_Duration_Seconds'].sum()
                total_input_hours = total_input_time / 3600
                
                # Time when input operators >= required
                input_available_mask = input_data['Person_Count_Imputed'] >= self.config.required_input_operators
                input_available_time = input_data.loc[input_available_mask, 'Adjusted_Duration_Seconds'].sum()
                input_available_hours = input_available_time / 3600
                
                # Input shortage time
                input_shortage_mask = input_data['Person_Count_Imputed'] < self.config.required_input_operators
                input_shortage_time = input_data.loc[input_shortage_mask, 'Adjusted_Duration_Seconds'].sum()
                input_shortage_hours = input_shortage_time / 3600
                
                results[shift]['Input_Availability_Hours'] = round(input_available_hours, 2)
                results[shift]['Input_Shortage_Hours'] = round(input_shortage_hours, 2)
                results[shift]['Input_Availability_Percent'] = round(
                    (input_available_hours / total_input_hours * 100) if total_input_hours > 0 else 0, 2
                )
            else:
                results[shift]['Input_Availability_Hours'] = 0
                results[shift]['Input_Shortage_Hours'] = 0
                results[shift]['Input_Availability_Percent'] = 0
            
            # Calculate output side operator availability
            output_mask = shift_data['Side'] == 'Output'
            output_data = shift_data[output_mask].copy()
            
            if len(output_data) > 0:
                total_output_time = output_data['Adjusted_Duration_Seconds'].sum()
                total_output_hours = total_output_time / 3600
                
                output_available_mask = output_data['Person_Count_Imputed'] >= self.config.required_output_operators
                output_available_time = output_data.loc[output_available_mask, 'Adjusted_Duration_Seconds'].sum()
                output_available_hours = output_available_time / 3600
                
                output_shortage_mask = output_data['Person_Count_Imputed'] < self.config.required_output_operators
                output_shortage_time = output_data.loc[output_shortage_mask, 'Adjusted_Duration_Seconds'].sum()
                output_shortage_hours = output_shortage_time / 3600
                
                results[shift]['Output_Availability_Hours'] = round(output_available_hours, 2)
                results[shift]['Output_Shortage_Hours'] = round(output_shortage_hours, 2)
                results[shift]['Output_Availability_Percent'] = round(
                    (output_available_hours / total_output_hours * 100) if total_output_hours > 0 else 0, 2
                )
            else:
                results[shift]['Output_Availability_Hours'] = 0
                results[shift]['Output_Shortage_Hours'] = 0
                results[shift]['Output_Availability_Percent'] = 0
        
        self.operator_kpis = results
        logger.info("Operator KPIs calculated successfully")
        return results
    
    def calculate_material_kpis(self) -> Dict[str, Any]:
        """
        Calculate material flow KPIs.
        
        Returns:
            Dictionary with material KPI results
        """
        logger.info("Calculating material KPIs...")
        
        results = {}
        
        for shift in ['Shift 1', 'Shift 2', 'Shift 3']:
            shift_data = self.processed_df[self.processed_df['Shift'] == shift].copy()
            
            if len(shift_data) == 0:
                results[shift] = {
                    'Processing_Hours': 0,
                    'Shortage_Hours': 0,
                    'Interruption_Hours': 0,
                    'Handling_Efficiency_Percent': 0
                }
                continue
            
            results[shift] = {}
            
            # Material processing duration (Box Present and machine working)
            processing_mask = (shift_data['Has_Box_Imputed'] == True) & (shift_data['Machine_Status_Imputed'] == 'Working')
            processing_time = shift_data.loc[processing_mask, 'Adjusted_Duration_Seconds'].sum()
            processing_hours = processing_time / 3600
            
            # Material shortage duration (No Box Present)
            shortage_mask = shift_data['Has_No_Box_Imputed'] == True
            shortage_time = shift_data.loc[shortage_mask, 'Adjusted_Duration_Seconds'].sum()
            shortage_hours = shortage_time / 3600
            
            # Material flow interruption (No Box and machine not working)
            interruption_mask = (shift_data['Has_No_Box_Imputed'] == True) & (shift_data['Machine_Status_Imputed'] == 'Not Working')
            interruption_time = shift_data.loc[interruption_mask, 'Adjusted_Duration_Seconds'].sum()
            interruption_hours = interruption_time / 3600
            
            # Material handling efficiency
            feeding_mask = shift_data['Is_Feeding'] == True
            feeding_time = shift_data.loc[feeding_mask, 'Adjusted_Duration_Seconds'].sum()
            feeding_hours = feeding_time / 3600
            
            bringing_mask = shift_data['Is_Bringing_Material'] == True
            bringing_time = shift_data.loc[bringing_mask, 'Adjusted_Duration_Seconds'].sum()
            bringing_hours = bringing_time / 3600
            
            handling_hours = feeding_hours + bringing_hours
            
            efficiency = (processing_hours / (processing_hours + handling_hours) * 100) if (processing_hours + handling_hours) > 0 else 0
            
            results[shift]['Processing_Hours'] = round(processing_hours, 2)
            results[shift]['Shortage_Hours'] = round(shortage_hours, 2)
            results[shift]['Interruption_Hours'] = round(interruption_hours, 2)
            results[shift]['Handling_Efficiency_Percent'] = round(efficiency, 2)
        
        self.material_kpis = results
        logger.info("Material KPIs calculated successfully")
        return results
    
    def generate_exception_report(self) -> pd.DataFrame:
        """
        Generate exception report with all anomaly events.
        
        Returns:
            DataFrame with exception report
        """
        logger.info("Generating exception report...")
        
        exceptions = []
        
        # Get all events with adjusted duration > 0
        events = self.processed_df[self.processed_df['Adjusted_Duration_Seconds'] > 0].copy()
        
        for idx, row in events.iterrows():
            duration_minutes = row['Adjusted_Duration_Seconds'] / 60
            
            # Machine downtime events
            if row['Machine_Status_Imputed'] == 'Not Working' and duration_minutes >= 1:
                exceptions.append({
                    'Timestamp': row['Timestamp'],
                    'Machine_ID': row['Machine_ID'],
                    'Shift': row['Shift'],
                    'Side': row['Side'],
                    'Exception_Type': 'Machine_Downtime',
                    'Description': f"Machine {row['Machine_ID']} not working",
                    'Duration_Seconds': row['Adjusted_Duration_Seconds'],
                    'Duration_Minutes': round(duration_minutes, 2),
                    'Severity': 'High' if duration_minutes > 30 else 'Medium'
                })
            
            # Operator shortage events
            if row['Side'] == 'Input' and row['Person_Count_Imputed'] < self.config.required_input_operators:
                if duration_minutes >= 1:
                    exceptions.append({
                        'Timestamp': row['Timestamp'],
                        'Machine_ID': row['Machine_ID'],
                        'Shift': row['Shift'],
                        'Side': 'Input',
                        'Exception_Type': 'Operator_Shortage',
                        'Description': f"Input operators: {row['Person_Count_Imputed']} (Required: {self.config.required_input_operators})",
                        'Duration_Seconds': row['Adjusted_Duration_Seconds'],
                        'Duration_Minutes': round(duration_minutes, 2),
                        'Severity': 'High' if duration_minutes > 30 else 'Medium'
                    })
            
            if row['Side'] == 'Output' and row['Person_Count_Imputed'] < self.config.required_output_operators:
                if duration_minutes >= 1:
                    exceptions.append({
                        'Timestamp': row['Timestamp'],
                        'Machine_ID': row['Machine_ID'],
                        'Shift': row['Shift'],
                        'Side': 'Output',
                        'Exception_Type': 'Operator_Shortage',
                        'Description': f"Output operators: {row['Person_Count_Imputed']} (Required: {self.config.required_output_operators})",
                        'Duration_Seconds': row['Adjusted_Duration_Seconds'],
                        'Duration_Minutes': round(duration_minutes, 2),
                        'Severity': 'High' if duration_minutes > 30 else 'Medium'
                    })
            
            # Material shortage events
            if row['Has_No_Box_Imputed'] and duration_minutes >= 1:
                exceptions.append({
                    'Timestamp': row['Timestamp'],
                    'Machine_ID': row['Machine_ID'],
                    'Shift': row['Shift'],
                    'Side': row['Side'],
                    'Exception_Type': 'Material_Shortage',
                    'Description': "No box present",
                    'Duration_Seconds': row['Adjusted_Duration_Seconds'],
                    'Duration_Minutes': round(duration_minutes, 2),
                    'Severity': 'High' if duration_minutes > 30 else 'Medium'
                })
            
            # Material flow interruption
            if row['Has_No_Box_Imputed'] and row['Machine_Status_Imputed'] == 'Not Working':
                if duration_minutes >= 1:
                    exceptions.append({
                        'Timestamp': row['Timestamp'],
                        'Machine_ID': row['Machine_ID'],
                        'Shift': row['Shift'],
                        'Side': row['Side'],
                        'Exception_Type': 'Flow_Interruption',
                        'Description': "No box and machine not working",
                        'Duration_Seconds': row['Adjusted_Duration_Seconds'],
                        'Duration_Minutes': round(duration_minutes, 2),
                        'Severity': 'High' if duration_minutes > 30 else 'Medium'
                    })
        
        self.exception_report = pd.DataFrame(exceptions)
        
        if len(self.exception_report) > 0:
            self.exception_report = self.exception_report.sort_values('Timestamp').reset_index(drop=True)
            
            # Format timestamp
            self.exception_report['Timestamp'] = self.exception_report['Timestamp'].dt.strftime('%Y-%m-%d %H:%M:%S')
        
        logger.info(f"Exception report generated: {len(self.exception_report)} exceptions")
        return self.exception_report
    
    def calculate_shift_kpis(self) -> pd.DataFrame:
        """
        Aggregate KPIs by shift.
        
        Returns:
            DataFrame with shift-wise KPIs
        """
        logger.info("Calculating shift-wise KPIs...")
        
        kpi_data = []
        
        for shift in ['Shift 1', 'Shift 2', 'Shift 3']:
            row = {'Shift': shift}
            
            # Machine KPIs (average across machines)
            machine_total_running = 0
            machine_total_downtime = 0
            machine_total_time = 0
            machine_count = 0
            
            for machine_id in self.machine_kpis.keys():
                if shift in self.machine_kpis[machine_id]:
                    mk = self.machine_kpis[machine_id][shift]
                    machine_total_running += mk['Running_Time_Hours']
                    machine_total_downtime += mk['Downtime_Hours']
                    machine_total_time += mk['Total_Time_Hours']
                    machine_count += 1
            
            if machine_count > 0:
                row['Avg_Running_Time_Hours'] = round(machine_total_running / machine_count, 2)
                row['Avg_Downtime_Hours'] = round(machine_total_downtime / machine_count, 2)
                row['Avg_Utilization_Percent'] = round(
                    (machine_total_running / machine_total_time * 100) if machine_total_time > 0 else 0, 2
                )
                row['Avg_Availability_Percent'] = round(
                    (machine_total_running / (machine_total_running + machine_total_downtime) * 100) 
                    if (machine_total_running + machine_total_downtime) > 0 else 0, 2
                )
            else:
                row['Avg_Running_Time_Hours'] = 0
                row['Avg_Downtime_Hours'] = 0
                row['Avg_Utilization_Percent'] = 0
                row['Avg_Availability_Percent'] = 0
            
            # Operator KPIs
            if shift in self.operator_kpis:
                ok = self.operator_kpis[shift]
                row['Input_Availability_Percent'] = ok['Input_Availability_Percent']
                row['Output_Availability_Percent'] = ok['Output_Availability_Percent']
                row['Input_Shortage_Hours'] = ok['Input_Shortage_Hours']
                row['Output_Shortage_Hours'] = ok['Output_Shortage_Hours']
            else:
                row['Input_Availability_Percent'] = 0
                row['Output_Availability_Percent'] = 0
                row['Input_Shortage_Hours'] = 0
                row['Output_Shortage_Hours'] = 0
            
            # Material KPIs
            if shift in self.material_kpis:
                mk = self.material_kpis[shift]
                row['Material_Processing_Hours'] = mk['Processing_Hours']
                row['Material_Shortage_Hours'] = mk['Shortage_Hours']
                row['Material_Interruption_Hours'] = mk['Interruption_Hours']
                row['Handling_Efficiency_Percent'] = mk['Handling_Efficiency_Percent']
            else:
                row['Material_Processing_Hours'] = 0
                row['Material_Shortage_Hours'] = 0
                row['Material_Interruption_Hours'] = 0
                row['Handling_Efficiency_Percent'] = 0
            
            kpi_data.append(row)
        
        self.shift_kpis = pd.DataFrame(kpi_data)
        logger.info(f"Shift KPIs calculated: {len(self.shift_kpis)} shifts")
        return self.shift_kpis
    
    def generate_kpi_summary(self) -> pd.DataFrame:
        """
        Generate comprehensive KPI summary.
        
        Returns:
            DataFrame with KPI summary
        """
        logger.info("Generating KPI summary...")
        
        summary_rows = []
        
        # Machine KPIs by machine and shift
        for machine_id in self.machine_kpis.keys():
            for shift in ['Shift 1', 'Shift 2', 'Shift 3']:
                if shift in self.machine_kpis[machine_id]:
                    mk = self.machine_kpis[machine_id][shift]
                    summary_rows.append({
                        'Category': 'Machine',
                        'Machine_ID': machine_id,
                        'Shift': shift,
                        'KPI': 'Running Time (Hours)',
                        'Value': mk['Running_Time_Hours'],
                        'Unit': 'Hours'
                    })
                    summary_rows.append({
                        'Category': 'Machine',
                        'Machine_ID': machine_id,
                        'Shift': shift,
                        'KPI': 'Downtime (Hours)',
                        'Value': mk['Downtime_Hours'],
                        'Unit': 'Hours'
                    })
                    summary_rows.append({
                        'Category': 'Machine',
                        'Machine_ID': machine_id,
                        'Shift': shift,
                        'KPI': 'Utilization %',
                        'Value': mk['Utilization_Percent'],
                        'Unit': '%'
                    })
                    summary_rows.append({
                        'Category': 'Machine',
                        'Machine_ID': machine_id,
                        'Shift': shift,
                        'KPI': 'Availability %',
                        'Value': mk['Availability_Percent'],
                        'Unit': '%'
                    })
        
        # Operator KPIs by shift
        for shift in ['Shift 1', 'Shift 2', 'Shift 3']:
            if shift in self.operator_kpis:
                ok = self.operator_kpis[shift]
                summary_rows.append({
                    'Category': 'Operator',
                    'Machine_ID': 'N/A',
                    'Shift': shift,
                    'KPI': 'Input Availability %',
                    'Value': ok['Input_Availability_Percent'],
                    'Unit': '%'
                })
                summary_rows.append({
                    'Category': 'Operator',
                    'Machine_ID': 'N/A',
                    'Shift': shift,
                    'KPI': 'Output Availability %',
                    'Value': ok['Output_Availability_Percent'],
                    'Unit': '%'
                })
                summary_rows.append({
                    'Category': 'Operator',
                    'Machine_ID': 'N/A',
                    'Shift': shift,
                    'KPI': 'Input Shortage (Hours)',
                    'Value': ok['Input_Shortage_Hours'],
                    'Unit': 'Hours'
                })
                summary_rows.append({
                    'Category': 'Operator',
                    'Machine_ID': 'N/A',
                    'Shift': shift,
                    'KPI': 'Output Shortage (Hours)',
                    'Value': ok['Output_Shortage_Hours'],
                    'Unit': 'Hours'
                })
        
        # Material KPIs by shift
        for shift in ['Shift 1', 'Shift 2', 'Shift 3']:
            if shift in self.material_kpis:
                mk = self.material_kpis[shift]
                summary_rows.append({
                    'Category': 'Material',
                    'Machine_ID': 'N/A',
                    'Shift': shift,
                    'KPI': 'Processing (Hours)',
                    'Value': mk['Processing_Hours'],
                    'Unit': 'Hours'
                })
                summary_rows.append({
                    'Category': 'Material',
                    'Machine_ID': 'N/A',
                    'Shift': shift,
                    'KPI': 'Shortage (Hours)',
                    'Value': mk['Shortage_Hours'],
                    'Unit': 'Hours'
                })
                summary_rows.append({
                    'Category': 'Material',
                    'Machine_ID': 'N/A',
                    'Shift': shift,
                    'KPI': 'Interruption (Hours)',
                    'Value': mk['Interruption_Hours'],
                    'Unit': 'Hours'
                })
                summary_rows.append({
                    'Category': 'Material',
                    'Machine_ID': 'N/A',
                    'Shift': shift,
                    'KPI': 'Handling Efficiency %',
                    'Value': mk['Handling_Efficiency_Percent'],
                    'Unit': '%'
                })
        
        self.kpi_summary = pd.DataFrame(summary_rows)
        logger.info(f"KPI summary generated: {len(self.kpi_summary)} records")
        return self.kpi_summary
    
    def generate_management_summary(self) -> pd.DataFrame:
        """
        Generate management summary with key insights.
        
        Returns:
            DataFrame with management summary
        """
        logger.info("Generating management summary...")
        
        summary_data = []
        
        # Overall statistics
        total_events = len(self.processed_df)
        total_duration = self.processed_df['Adjusted_Duration_Seconds'].sum() / 3600
        
        summary_data.append({
            'Metric': 'Total Events Processed',
            'Value': total_events,
            'Unit': 'Events',
            'Insight': 'Total number of valid events analyzed'
        })
        
        summary_data.append({
            'Metric': 'Total Duration',
            'Value': round(total_duration, 2),
            'Unit': 'Hours',
            'Insight': 'Total adjusted time across all events'
        })
        
        # Machine statistics
        total_running = 0
        total_downtime = 0
        for machine_id in self.machine_kpis.keys():
            for shift in ['Shift 1', 'Shift 2', 'Shift 3']:
                if shift in self.machine_kpis[machine_id]:
                    mk = self.machine_kpis[machine_id][shift]
                    total_running += mk['Running_Time_Hours']
                    total_downtime += mk['Downtime_Hours']
        
        summary_data.append({
            'Metric': 'Total Running Time',
            'Value': round(total_running, 2),
            'Unit': 'Hours',
            'Insight': 'Total machine running time across all shifts'
        })
        
        summary_data.append({
            'Metric': 'Total Downtime',
            'Value': round(total_downtime, 2),
            'Unit': 'Hours',
            'Insight': 'Total machine downtime across all shifts'
        })
        
        overall_utilization = (total_running / (total_running + total_downtime) * 100) if (total_running + total_downtime) > 0 else 0
        summary_data.append({
            'Metric': 'Overall Utilization',
            'Value': round(overall_utilization, 2),
            'Unit': '%',
            'Insight': 'Overall machine utilization rate'
        })
        
        # Operator statistics
        total_input_shortage = 0
        total_output_shortage = 0
        for shift in ['Shift 1', 'Shift 2', 'Shift 3']:
            if shift in self.operator_kpis:
                ok = self.operator_kpis[shift]
                total_input_shortage += ok['Input_Shortage_Hours']
                total_output_shortage += ok['Output_Shortage_Hours']
        
        summary_data.append({
            'Metric': 'Total Input Shortage',
            'Value': round(total_input_shortage, 2),
            'Unit': 'Hours',
            'Insight': 'Total input operator shortage time'
        })
        
        summary_data.append({
            'Metric': 'Total Output Shortage',
            'Value': round(total_output_shortage, 2),
            'Unit': 'Hours',
            'Insight': 'Total output operator shortage time'
        })
        
        # Material statistics
        total_material_shortage = 0
        for shift in ['Shift 1', 'Shift 2', 'Shift 3']:
            if shift in self.material_kpis:
                mk = self.material_kpis[shift]
                total_material_shortage += mk['Shortage_Hours']
        
        summary_data.append({
            'Metric': 'Total Material Shortage',
            'Value': round(total_material_shortage, 2),
            'Unit': 'Hours',
            'Insight': 'Total material shortage time'
        })
        
        # Exception summary
        if self.exception_report is not None and len(self.exception_report) > 0:
            exception_counts = self.exception_report['Exception_Type'].value_counts().to_dict()
            for ex_type, count in exception_counts.items():
                summary_data.append({
                    'Metric': f'{ex_type} Events',
                    'Value': count,
                    'Unit': 'Events',
                    'Insight': f'Number of {ex_type} events detected'
                })
        
        self.management_summary = pd.DataFrame(summary_data)
        logger.info(f"Management summary generated: {len(self.management_summary)} records")
        return self.management_summary
    
    def generate_calculation_audit(self) -> pd.DataFrame:
        """
        Generate calculation audit trail.
        
        Returns:
            DataFrame with calculation audit records
        """
        logger.info("Generating calculation audit...")
        
        audit_records = []
        
        # Record data loading
        audit_records.append({
            'Step': 'Data Loading',
            'Description': 'Load processed data from script1 output',
            'Value': f"{len(self.processed_df)} records loaded",
            'Formula': 'N/A',
            'Source': 'processed_data_script1_final.xlsx',
            'Timestamp': datetime.now().isoformat()
        })
        
        # Record machine KPI calculation
        audit_records.append({
            'Step': 'Machine KPI Calculation',
            'Description': 'Calculate running time, downtime, utilization, availability',
            'Value': f"{len(self.machine_kpis)} machines analyzed",
            'Formula': 'Sum(Adjusted_Duration_Seconds) where Machine_Status = Working/Not Working',
            'Source': 'Processed Data',
            'Timestamp': datetime.now().isoformat()
        })
        
        # Record operator KPI calculation
        audit_records.append({
            'Step': 'Operator KPI Calculation',
            'Description': 'Calculate operator availability and shortage',
            'Value': f"Required: {self.config.required_input_operators} input, {self.config.required_output_operators} output",
            'Formula': 'Sum(Adjusted_Duration_Seconds) where Person_Count >= Required',
            'Source': 'Processed Data',
            'Timestamp': datetime.now().isoformat()
        })
        
        # Record material KPI calculation
        audit_records.append({
            'Step': 'Material KPI Calculation',
            'Description': 'Calculate processing, shortage, interruption, handling efficiency',
            'Value': f"{len(self.material_kpis)} shifts analyzed",
            'Formula': 'Sum(Adjusted_Duration_Seconds) based on Box status and Feeding status',
            'Source': 'Processed Data',
            'Timestamp': datetime.now().isoformat()
        })
        
        # Record exception report generation
        exception_count = len(self.exception_report) if self.exception_report is not None else 0
        audit_records.append({
            'Step': 'Exception Report Generation',
            'Description': 'Identify machine downtime, operator shortage, material shortage, flow interruption',
            'Value': f"{exception_count} exceptions found",
            'Formula': 'Filter events where status indicates exception',
            'Source': 'Processed Data',
            'Timestamp': datetime.now().isoformat()
        })
        
        audit_df = pd.DataFrame(audit_records)
        self.calculation_audit = audit_df
        logger.info(f"Calculation audit generated: {len(audit_df)} records")
        return audit_df
    
    def generate_top_productivity_loss_events(self) -> pd.DataFrame:
        """
        Identify top 5 productivity loss events.
        
        Returns:
            DataFrame with top productivity loss events
        """
        logger.info("Generating top productivity loss events...")
        
        if self.exception_report is None or len(self.exception_report) == 0:
            logger.info("No exceptions found for productivity loss analysis")
            return pd.DataFrame()
        
        # Group by exception type and sum durations
        loss_data = self.exception_report.copy()
        loss_data['Duration_Seconds'] = pd.to_numeric(loss_data['Duration_Seconds'], errors='coerce')
        
        # Group by type and get top 5
        top_losses = (
            loss_data.groupby('Exception_Type')
            .agg({
                'Duration_Seconds': 'sum',
                'Duration_Minutes': 'sum'
            })
            .reset_index()
            .sort_values('Duration_Seconds', ascending=False)
            .head(5)
        )
        
        top_losses['Duration_Hours'] = top_losses['Duration_Seconds'] / 3600
        
        top_losses = top_losses.rename(columns={
            'Exception_Type': 'Loss_Event_Type',
            'Duration_Seconds': 'Total_Duration_Seconds',
            'Duration_Minutes': 'Total_Duration_Minutes',
            'Duration_Hours': 'Total_Duration_Hours'
        })
        
        top_losses = top_losses[['Loss_Event_Type', 'Total_Duration_Seconds', 'Total_Duration_Minutes', 'Total_Duration_Hours']]
        top_losses['Total_Duration_Hours'] = top_losses['Total_Duration_Hours'].round(2)
        top_losses['Total_Duration_Minutes'] = top_losses['Total_Duration_Minutes'].round(2)
        
        logger.info(f"Top 5 productivity loss events identified")
        return top_losses
    
    def save_to_excel(self, output_file: str = 'kpi_output_script2_final.xlsx') -> None:
        """
        Save all KPI results to Excel workbook.
        
        Args:
            output_file: Output Excel file path
        """
        logger.info(f"Saving KPI results to Excel: {output_file}")
        
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Sheet 1: KPI Summary
                if self.kpi_summary is not None and not self.kpi_summary.empty:
                    self.kpi_summary.to_excel(writer, sheet_name='KPI Summary', index=False)
                
                # Sheet 2: Shift KPIs
                if self.shift_kpis is not None and not self.shift_kpis.empty:
                    self.shift_kpis.to_excel(writer, sheet_name='Shift KPIs', index=False)
                
                # Sheet 3: Exception Report
                if self.exception_report is not None and not self.exception_report.empty:
                    self.exception_report.to_excel(writer, sheet_name='Exception Report', index=False)
                
                # Sheet 4: Top Productivity Loss
                top_losses = self.generate_top_productivity_loss_events()
                if not top_losses.empty:
                    top_losses.to_excel(writer, sheet_name='Top 5 Loss Events', index=False)
                
                # Sheet 5: Management Summary
                if self.management_summary is not None and not self.management_summary.empty:
                    self.management_summary.to_excel(writer, sheet_name='Management Summary', index=False)
                
                # Sheet 6: Calculation Audit
                if self.calculation_audit is not None and not self.calculation_audit.empty:
                    self.calculation_audit.to_excel(writer, sheet_name='Calculation Audit', index=False)
                
                # Auto-adjust column widths
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if cell.value is not None and len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                
                logger.info(f"Excel file saved successfully: {output_file}")
                
        except Exception as e:
            logger.error(f"Error saving Excel file: {str(e)}", exc_info=True)
            raise
    
    def run(self, input_file: str, output_file: str) -> None:
        """
        Execute the complete KPI calculation pipeline.
        
        Args:
            input_file: Input processed data file
            output_file: Output KPI results file
        """
        logger.info("=" * 80)
        logger.info("KPI ENGINE - SCRIPT 2 OF PROJECT")
        logger.info("=" * 80)
        
        try:
            # Step 1: Load processed data
            self.load_processed_data(input_file)
            
            # Step 2: Calculate Machine KPIs
            self.calculate_machine_kpis()
            
            # Step 3: Calculate Operator KPIs
            self.calculate_operator_kpis()
            
            # Step 4: Calculate Material KPIs
            self.calculate_material_kpis()
            
            # Step 5: Generate Exception Report
            self.generate_exception_report()
            
            # Step 6: Calculate Shift KPIs
            self.calculate_shift_kpis()
            
            # Step 7: Generate KPI Summary
            self.generate_kpi_summary()
            
            # Step 8: Generate Management Summary
            self.generate_management_summary()
            
            # Step 9: Generate Calculation Audit
            self.generate_calculation_audit()
            
            # Step 10: Save to Excel
            self.save_to_excel(output_file)
            
            logger.info("=" * 80)
            logger.info("KPI ENGINE EXECUTION COMPLETED SUCCESSFULLY")
            logger.info(f"OUTPUT FILE: {output_file}")
            logger.info("=" * 80)
            
            # Print summary statistics
            logger.info("\nSUMMARY STATISTICS:")
            if self.processed_df is not None:
                logger.info(f"Total records analyzed: {len(self.processed_df)}")
            
            if self.machine_kpis:
                logger.info(f"Machines analyzed: {len(self.machine_kpis)}")
            
            if self.exception_report is not None:
                logger.info(f"Exceptions found: {len(self.exception_report)}")
                
                if len(self.exception_report) > 0:
                    exception_counts = self.exception_report['Exception_Type'].value_counts()
                    for ex_type, count in exception_counts.items():
                        logger.info(f"  {ex_type}: {count}")
            
            if self.kpi_summary is not None:
                logger.info(f"KPI Summary records: {len(self.kpi_summary)}")
            
            if self.management_summary is not None:
                logger.info(f"Management Summary records: {len(self.management_summary)}")
            
        except Exception as e:
            logger.error(f"Critical error in KPI engine: {str(e)}", exc_info=True)
            raise


def main():
    """
    Main execution function.
    """
    # Define file paths
    input_file = 'processed_data_script1_final.xlsx'
    output_file = 'kpi_output_script2_final.xlsx'
    
    # Check if input file exists
    input_path = Path(input_file)
    if not input_path.exists():
        logger.error(f"Input file not found: {input_file}")
        logger.info("Please run script1 first to generate the processed data")
        return
    
    # Initialize configuration
    config = KPIEngineConfig()
    logger.info("Configuration initialized successfully")
    
    # Initialize and run KPI engine
    engine = KPIEngine(config)
    engine.run(input_file, output_file)


if __name__ == "__main__":
    main