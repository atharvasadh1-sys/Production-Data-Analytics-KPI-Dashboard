#!/usr/bin/env python3
"""
Machines Input Side Data 22 May 2026 to 23 May 2026.xlsx
Machines Output Side Data 22 May 2026 to 23 May 2026.xlsx
Assignment for shortlisted candidates - Round 1

This script processes the raw data files, performs comprehensive data cleaning,
event parsing, shift assignment, and duration calculations for all machine events.

Author: Senior Python Data Engineer
Date: 2026-08-04
Version: 2.0.0 - Pandas 2.x Compatible
"""

import pandas as pd
import numpy as np
import re
import logging
from datetime import datetime, timedelta, time
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from typing import Dict, List, Tuple, Optional, Any, Union
import warnings
warnings.filterwarnings('ignore')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('data_processing.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class DataProcessingConfig:
    """
    Configuration class for data processing parameters.
    All configurable values are centralized here for easy modification.
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
        
        # Noise filtering thresholds
        self.min_event_duration_seconds = 30
        self.max_gap_for_imputation_seconds = 300  # 5 minutes
        
        # Duplicate handling
        self.duplicate_tolerance_seconds = 2  # Events within 2 seconds are considered duplicates
        
        # Date range for data (from file names)
        self.data_start_date = datetime(2026, 5, 22, 0, 0, 0)
        self.data_end_date = datetime(2026, 5, 23, 23, 59, 59)


class DataProcessor:
    """
    Main data processing class that handles the complete data pipeline.
    """
    
    def __init__(self, config: DataProcessingConfig):
        self.config = config
        self.input_df = None
        self.output_df = None
        self.processed_df = None
        self.invalid_df = None
        self.duplicate_df = None
        self.noise_df = None
        self.calculation_audit = None
        
    def load_data(self, input_file: str, output_file: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        Load input and output data files with proper column handling.
        
        Args:
            input_file: Path to input side data file
            output_file: Path to output side data file
            
        Returns:
            Tuple of (input_df, output_df)
        """
        logger.info(f"Loading input data from: {input_file}")
        logger.info(f"Loading output data from: {output_file}")
        
        try:
            # Load input data
            self.input_df = pd.read_excel(input_file)
            logger.info(f"Input data loaded: {len(self.input_df)} rows")
            
            # Load output data
            self.output_df = pd.read_excel(output_file)
            logger.info(f"Output data loaded: {len(self.output_df)} rows")
            
            # Validate columns
            required_cols = ['Created At', 'Content', 'Video Link']
            for col in required_cols:
                if col not in self.input_df.columns:
                    raise KeyError(f"Required column '{col}' missing from input data")
                if col not in self.output_df.columns:
                    raise KeyError(f"Required column '{col}' missing from output data")
            
            # Rename columns to avoid space issues
            self.input_df.rename(columns={
                'Created At': 'Created_At',
                'Video Link': 'Video_Link'
            }, inplace=True)
            self.output_df.rename(columns={
                'Created At': 'Created_At',
                'Video Link': 'Video_Link'
            }, inplace=True)
            
            return self.input_df, self.output_df
            
        except Exception as e:
            logger.error(f"Error loading data: {str(e)}")
            raise
    
    def parse_timestamp(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Parse timestamps and create datetime objects with proper timezone handling.
        
        Args:
            df: DataFrame with 'Created_At' column
            
        Returns:
            DataFrame with parsed timestamps
        """
        logger.info("Parsing timestamps...")
        
        # Create a copy to avoid modifying original
        df_parsed = df.copy()
        
        # Parse timestamps
        df_parsed['Timestamp'] = pd.to_datetime(df_parsed['Created_At'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
        
        # Drop rows with invalid timestamps
        invalid_timestamps = df_parsed['Timestamp'].isna().sum()
        if invalid_timestamps > 0:
            logger.warning(f"Found {invalid_timestamps} rows with invalid timestamps. Dropping them.")
            df_parsed = df_parsed.dropna(subset=['Timestamp'])
        
        # Extract date and time components
        df_parsed['Date'] = df_parsed['Timestamp'].dt.date
        df_parsed['Time'] = df_parsed['Timestamp'].dt.time
        df_parsed['Hour'] = df_parsed['Timestamp'].dt.hour
        df_parsed['Minute'] = df_parsed['Timestamp'].dt.minute
        df_parsed['Second'] = df_parsed['Timestamp'].dt.second
        df_parsed['Day_of_Week'] = df_parsed['Timestamp'].dt.day_name()
        
        # Sort by timestamp
        df_parsed = df_parsed.sort_values('Timestamp').reset_index(drop=True)
        
        logger.info(f"Timestamp parsing complete: {len(df_parsed)} rows")
        return df_parsed
    
    def extract_machine_id(self, content: str) -> Optional[int]:
        """
        Extract machine ID (1 or 2) from content string.
        
        Args:
            content: Raw content string
            
        Returns:
            Machine ID or None if not found
        """
        if not isinstance(content, str):
            return None
            
        patterns = [
            r'Machine\s*([12])',
            r'[Mm]achine\s*#?\s*([12])',
            r'[Mm]achine\s*([12])\s*[Ii]nput',
            r'[Mm]achine\s*([12])\s*[Oo]utput'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, content)
            if match:
                return int(match.group(1))
        
        return None
    
    def extract_side(self, content: str) -> str:
        """
        Determine if event is from input or output side.
        
        Args:
            content: Raw content string
            
        Returns:
            'Input', 'Output', or 'Unknown'
        """
        if not isinstance(content, str):
            return 'Unknown'
        
        if 'Input' in content:
            return 'Input'
        elif 'Output' in content:
            return 'Output'
        else:
            return 'Unknown'
    
    def extract_person_count(self, content: str) -> Optional[int]:
        """
        Extract the number of persons from content string.
        
        Args:
            content: Raw content string
            
        Returns:
            Person count or None if not found
        """
        if not isinstance(content, str):
            return None
            
        patterns = [
            r'total number of person at Machine [12] (?:Input|Output) machine:?\s*(\d+)',
            r'total number of person at Machine [12] (?:Input|Output) machine\s*:\s*(\d+)',
            r'person(?:s)?\s*:\s*(\d+)',
            r'count\s*:\s*(\d+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, content)
            if match:
                return int(match.group(1))
        
        return None
    
    def extract_machine_status(self, content: str) -> str:
        """
        Extract machine working status from content.
        
        Args:
            content: Raw content string
            
        Returns:
            'Working', 'Not Working', or 'Unknown'
        """
        if not isinstance(content, str):
            return 'Unknown'
            
        content_lower = content.lower()
        
        if 'not working' in content_lower:
            return 'Not Working'
        elif 'working' in content_lower and 'not working' not in content_lower:
            return 'Working'
        else:
            return 'Unknown'
    
    def extract_process_statuses(self, content: str) -> List[str]:
        """
        Extract all process statuses from content within square brackets.
        
        Args:
            content: Raw content string
            
        Returns:
            List of process status strings
        """
        if not isinstance(content, str):
            return []
        
        pattern = r'\[([^\]]+)\]'
        matches = re.findall(pattern, content)
        
        statuses = []
        for match in matches:
            parts = [p.strip() for p in match.split(',')]
            statuses.extend(parts)
        
        return statuses
    
    def classify_event_type(self, content: str) -> str:
        """
        Classify the event type based on content.
        
        Args:
            content: Raw content string
            
        Returns:
            Event type classification
        """
        if not isinstance(content, str):
            return 'Unknown'
        
        content_lower = content.lower()
        
        if 'total number of person' in content_lower:
            return 'Personnel_Count'
        elif 'input process' in content_lower or 'output process' in content_lower:
            return 'Process_Status'
        elif 'not working' in content_lower:
            return 'Machine_Downtime'
        elif 'working' in content_lower and 'not working' not in content_lower:
            return 'Machine_Working'
        else:
            return 'Other'
    
    def determine_is_break_time(self, timestamp: datetime) -> bool:
        """
        Check if timestamp falls within any break period.
        
        Args:
            timestamp: Datetime object
            
        Returns:
            True if within break period, False otherwise
        """
        t = timestamp.time()
        
        shift1_break_start = self.config.shifts['Shift 1']['break_start']
        shift1_break_end = self.config.shifts['Shift 1']['break_end']
        if shift1_break_start <= t < shift1_break_end:
            return True
        
        shift2_break_start = self.config.shifts['Shift 2']['break_start']
        shift2_break_end = self.config.shifts['Shift 2']['break_end']
        if shift2_break_start <= t < shift2_break_end:
            return True
        
        shift3_break_start = self.config.shifts['Shift 3']['break_start']
        shift3_break_end = self.config.shifts['Shift 3']['break_end']
        if shift3_break_start <= t < shift3_break_end:
            return True
        
        return False
    
    def determine_shift(self, timestamp: datetime) -> str:
        """
        Determine which shift the timestamp belongs to.
        
        Args:
            timestamp: Datetime object
            
        Returns:
            Shift name string
        """
        t = timestamp.time()
        
        if t >= time(6, 0, 0) and t < time(14, 0, 0):
            return 'Shift 1'
        elif t >= time(14, 0, 0) and t < time(22, 0, 0):
            return 'Shift 2'
        elif t >= time(22, 0, 0) or t < time(6, 0, 0):
            return 'Shift 3'
        else:
            return 'Unknown'
    
    def assign_shift_date(self, timestamp: datetime) -> datetime:
        """
        For Shift 3, determine the appropriate date to associate the shift with.
        
        Args:
            timestamp: Datetime object
            
        Returns:
            Shift date (Date that the shift belongs to)
        """
        shift_name = self.determine_shift(timestamp)
        
        if shift_name == 'Shift 3':
            if timestamp.hour < 6:
                return timestamp - timedelta(days=1)
        
        return timestamp
    
    def calculate_duration(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Calculate event durations based on timestamp differences.
        
        Args:
            df: DataFrame with parsed timestamps
            
        Returns:
            DataFrame with duration columns added
        """
        logger.info("Calculating event durations...")
        
        df_duration = df.copy()
        
        if 'Machine_ID' not in df_duration.columns:
            logger.warning("Machine_ID column missing. Adding default value.")
            df_duration['Machine_ID'] = 1
        
        df_duration = df_duration.sort_values(['Machine_ID', 'Timestamp']).reset_index(drop=True)
        
        df_duration['Next_Timestamp'] = df_duration.groupby('Machine_ID')['Timestamp'].shift(-1)
        df_duration['Duration_Seconds'] = (
            df_duration['Next_Timestamp'] - df_duration['Timestamp']
        ).dt.total_seconds()
        
        df_duration['Duration_Seconds'] = df_duration['Duration_Seconds'].fillna(0)
        df_duration['Duration_Seconds'] = df_duration['Duration_Seconds'].clip(lower=0)
        
        df_duration['Break_Overlap_Seconds'] = df_duration.apply(
            lambda row: self.calculate_break_overlap(
                row['Timestamp'],
                row['Duration_Seconds'],
                row['Next_Timestamp'] if pd.notna(row['Next_Timestamp']) else None
            ),
            axis=1
        )
        
        df_duration['Adjusted_Duration_Seconds'] = (
            df_duration['Duration_Seconds'] - df_duration['Break_Overlap_Seconds']
        )
        df_duration['Adjusted_Duration_Seconds'] = df_duration['Adjusted_Duration_Seconds'].clip(lower=0)
        
        df_duration['Duration_Minutes'] = df_duration['Duration_Seconds'] / 60
        df_duration['Adjusted_Duration_Minutes'] = df_duration['Adjusted_Duration_Seconds'] / 60
        
        logger.info(f"Duration calculation complete. Total duration: {df_duration['Duration_Seconds'].sum():.2f} seconds")
        logger.info(f"Break overlap total: {df_duration['Break_Overlap_Seconds'].sum():.2f} seconds")
        
        return df_duration
    
    def calculate_break_overlap(self, start_time: datetime, duration_seconds: float, end_time: Optional[datetime]) -> float:
        """
        Calculate the amount of time that overlaps with break periods.
        
        Args:
            start_time: Event start time
            duration_seconds: Event duration in seconds
            end_time: Event end time (or None if last event)
            
        Returns:
            Overlap duration in seconds
        """
        if duration_seconds == 0 or duration_seconds is None or pd.isna(duration_seconds):
            return 0.0
        
        total_overlap = 0.0
        end_time_calc = start_time + timedelta(seconds=duration_seconds)
        
        break_periods = [
            (self.config.shifts['Shift 1']['break_start'], self.config.shifts['Shift 1']['break_end']),
            (self.config.shifts['Shift 2']['break_start'], self.config.shifts['Shift 2']['break_end']),
            (self.config.shifts['Shift 3']['break_start'], self.config.shifts['Shift 3']['break_end'])
        ]
        
        for break_start, break_end in break_periods:
            if break_start > break_end:
                # Crosses midnight (02:30-03:00)
                break_start_dt = datetime.combine(start_time.date(), break_start)
                break_end_dt = datetime.combine(start_time.date(), break_end) + timedelta(days=1)
                
                overlap_start = max(start_time, break_start_dt)
                overlap_end = min(end_time_calc, break_end_dt)
                
                if overlap_start < overlap_end:
                    overlap = (overlap_end - overlap_start).total_seconds()
                    if overlap > 0:
                        total_overlap += overlap
                
                break_start_dt_prev = datetime.combine(start_time.date() - timedelta(days=1), break_start)
                break_end_dt_prev = datetime.combine(start_time.date() - timedelta(days=1), break_end) + timedelta(days=1)
                
                overlap_start_prev = max(start_time, break_start_dt_prev)
                overlap_end_prev = min(end_time_calc, break_end_dt_prev)
                
                if overlap_start_prev < overlap_end_prev:
                    overlap_prev = (overlap_end_prev - overlap_start_prev).total_seconds()
                    if overlap_prev > 0:
                        total_overlap += overlap_prev
            else:
                break_start_dt = datetime.combine(start_time.date(), break_start)
                break_end_dt = datetime.combine(start_time.date(), break_end)
                
                overlap_start = max(start_time, break_start_dt)
                overlap_end = min(end_time_calc, break_end_dt)
                
                if overlap_start < overlap_end:
                    overlap = (overlap_end - overlap_start).total_seconds()
                    if overlap > 0:
                        total_overlap += overlap
        
        return min(total_overlap, duration_seconds)
    
    def parse_content_structured(self, content: str) -> Dict[str, Any]:
        """
        Parse content into structured fields.
        
        Args:
            content: Raw content string
            
        Returns:
            Dictionary of structured fields
        """
        result = {
            'Machine_ID': self.extract_machine_id(content),
            'Side': self.extract_side(content),
            'Person_Count': self.extract_person_count(content),
            'Machine_Status': self.extract_machine_status(content),
            'Process_Statuses': self.extract_process_statuses(content),
            'Event_Type': self.classify_event_type(content),
            'Is_Personnel_Event': 'total number of person' in content.lower() if isinstance(content, str) else False,
            'Is_Process_Event': 'process' in content.lower() if isinstance(content, str) else False,
            'Is_Machine_Status_Event': ('working' in content.lower() or 'not working' in content.lower()) if isinstance(content, str) else False
        }
        
        process_statuses_lower = [p.lower() for p in result['Process_Statuses']]
        result['Has_Box'] = any('box present' in p for p in process_statuses_lower)
        result['Has_No_Box'] = any('no box present' in p for p in process_statuses_lower)
        result['Is_Feeding'] = any('feeding' in p for p in process_statuses_lower)
        result['Is_Bringing_Material'] = any('bringing new material' in p for p in process_statuses_lower)
        result['Operator_Present'] = any('operator is present' in p for p in process_statuses_lower)
        result['Operator_Not_Present'] = any('operator not present' in p for p in process_statuses_lower)
        result['Material_Present_Operator_Absent'] = any('material present & operator not present' in p for p in process_statuses_lower)
        result['Operator_Present_Material_Absent'] = any('operator present & material not present' in p for p in process_statuses_lower)
        
        return result
    
    def validate_event(self, row: Dict[str, Any]) -> Tuple[bool, str]:
        """
        Validate an event for completeness and quality.
        
        Args:
            row: Event data dictionary
            
        Returns:
            Tuple of (is_valid, reason)
        """
        if pd.isna(row.get('Timestamp')):
            return False, "Missing timestamp"
        
        if pd.isna(row.get('Content')) or str(row.get('Content', '')).strip() == '':
            return False, "Empty content"
        
        if row.get('Machine_ID') is None:
            return False, "Machine ID not found"
        
        if row.get('Event_Type') == 'Unknown' and 'total' not in str(row.get('Content', '')).lower():
            return False, "Unknown event type"
        
        if len(str(row.get('Content', ''))) < 5:
            return False, "Content too short"
        
        return True, "Valid event"
    
    def remove_duplicates(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        Remove duplicate events based on timestamp and content similarity.
        
        Args:
            df: DataFrame with parsed events
            
        Returns:
            Tuple of (deduplicated_df, duplicate_df)
        """
        logger.info("Removing duplicates...")
        
        df_copy = df.copy()
        
        # Round timestamp to nearest second for deduplication
        df_copy['Timestamp_Floor'] = df_copy['Timestamp'].dt.floor('s')
        
        duplicate_mask = df_copy.duplicated(
            subset=['Timestamp_Floor', 'Machine_ID', 'Side'],
            keep='first'
        )
        
        df_unique = df_copy[~duplicate_mask].copy()
        df_duplicates = df_copy[duplicate_mask].copy()
        
        df_unique = df_unique.drop(columns=['Timestamp_Floor'])
        df_duplicates = df_duplicates.drop(columns=['Timestamp_Floor'])
        
        logger.info(f"Removed {len(df_duplicates)} duplicate events")
        logger.info(f"Unique events: {len(df_unique)}")
        
        return df_unique, df_duplicates
    
    def remove_noise_events(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        Remove short-duration noise events based on configured threshold.
        
        Args:
            df: DataFrame with duration calculations
            
        Returns:
            Tuple of (filtered_df, noise_df)
        """
        logger.info("Removing noise events...")
        
        if 'Adjusted_Duration_Seconds' not in df.columns:
            logger.warning("Duration column not found. Skipping noise filtering.")
            return df, pd.DataFrame()
        
        df_copy = df.copy()
        
        # Filter out events with duration less than threshold
        threshold = self.config.min_event_duration_seconds
        
        # Keep events with adjusted duration >= threshold
        noise_mask = (df_copy['Adjusted_Duration_Seconds'] < threshold) & (df_copy['Adjusted_Duration_Seconds'] > 0)
        
        df_filtered = df_copy[~noise_mask].copy()
        df_noise = df_copy[noise_mask].copy()
        
        logger.info(f"Removed {len(df_noise)} noise events (duration < {threshold}s)")
        logger.info(f"Filtered events: {len(df_filtered)}")
        
        return df_filtered, df_noise
    
    def remove_invalid_events(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        Remove invalid events and create invalid events dataframe.
        
        Args:
            df: DataFrame with parsed events
            
        Returns:
            Tuple of (valid_df, invalid_df)
        """
        logger.info("Removing invalid events...")
        
        valid_rows = []
        invalid_rows = []
        
        for idx, row in df.iterrows():
            row_dict = row.to_dict()
            is_valid, reason = self.validate_event(row_dict)
            
            row_dict['Is_Valid'] = is_valid
            row_dict['Invalid_Reason'] = reason if not is_valid else ''
            
            if is_valid:
                valid_rows.append(row_dict)
            else:
                invalid_rows.append(row_dict)
        
        valid_df = pd.DataFrame(valid_rows)
        invalid_df = pd.DataFrame(invalid_rows)
        
        logger.info(f"Valid events: {len(valid_df)}")
        logger.info(f"Invalid events: {len(invalid_df)}")
        
        return valid_df, invalid_df
    
    def merge_input_output_data(self, input_df: pd.DataFrame, output_df: pd.DataFrame) -> pd.DataFrame:
        """
        Merge input and output data into a single structured dataset.
        
        Args:
            input_df: Processed input data
            output_df: Processed output data
            
        Returns:
            Merged DataFrame
        """
        logger.info("Merging input and output data...")
        
        input_df_copy = input_df.copy()
        output_df_copy = output_df.copy()
        input_df_copy['Source'] = 'Input'
        output_df_copy['Source'] = 'Output'
        
        combined_df = pd.concat([input_df_copy, output_df_copy], ignore_index=True)
        
        combined_df = combined_df.sort_values(['Timestamp', 'Machine_ID', 'Side']).reset_index(drop=True)
        
        logger.info(f"Combined data: {len(combined_df)} rows")
        
        return combined_df
    
    def impute_missing_states(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Impute missing states using forward fill for gaps under threshold.
        
        Args:
            df: DataFrame with events
            
        Returns:
            DataFrame with imputed values
        """
        logger.info("Imputing missing states...")
        
        df_imputed = df.copy()
        
        df_imputed = df_imputed.sort_values(['Machine_ID', 'Timestamp']).reset_index(drop=True)
        
        # Initialize imputation columns with original values
        df_imputed['Machine_Status_Imputed'] = df_imputed['Machine_Status']
        df_imputed['Person_Count_Imputed'] = df_imputed['Person_Count']
        df_imputed['Has_Box_Imputed'] = df_imputed['Has_Box']
        df_imputed['Has_No_Box_Imputed'] = df_imputed['Has_No_Box']
        
        # Forward fill for each machine using ffill() (pandas 2.x compatible)
        for machine_id in df_imputed['Machine_ID'].unique():
            machine_mask = df_imputed['Machine_ID'] == machine_id
            
            df_imputed.loc[machine_mask, 'Machine_Status_Imputed'] = (
                df_imputed.loc[machine_mask, 'Machine_Status_Imputed'].ffill()
            )
            
            for side in ['Input', 'Output']:
                side_mask = machine_mask & (df_imputed['Side'] == side)
                if side_mask.any():
                    df_imputed.loc[side_mask, 'Person_Count_Imputed'] = (
                        df_imputed.loc[side_mask, 'Person_Count_Imputed'].ffill()
                    )
            
            df_imputed.loc[machine_mask, 'Has_Box_Imputed'] = (
                df_imputed.loc[machine_mask, 'Has_Box_Imputed'].ffill()
            )
            df_imputed.loc[machine_mask, 'Has_No_Box_Imputed'] = (
                df_imputed.loc[machine_mask, 'Has_No_Box_Imputed'].ffill()
            )
        
        # Fill remaining NAs with defaults
        df_imputed['Machine_Status_Imputed'] = df_imputed['Machine_Status_Imputed'].fillna('Unknown')
        df_imputed['Person_Count_Imputed'] = df_imputed['Person_Count_Imputed'].fillna(0)
        df_imputed['Has_Box_Imputed'] = df_imputed['Has_Box_Imputed'].fillna(False)
        df_imputed['Has_No_Box_Imputed'] = df_imputed['Has_No_Box_Imputed'].fillna(False)
        
        df_imputed['Is_Imputed'] = False
        df_imputed.loc[
            (df_imputed['Machine_Status_Imputed'] != df_imputed['Machine_Status']) |
            (df_imputed['Person_Count_Imputed'] != df_imputed['Person_Count']),
            'Is_Imputed'
        ] = True
        
        df_imputed['Time_Gap_Seconds'] = df_imputed.groupby('Machine_ID')['Timestamp'].diff().dt.total_seconds()
        
        df_imputed.loc[
            df_imputed['Time_Gap_Seconds'] > self.config.max_gap_for_imputation_seconds,
            'Is_Imputed'
        ] = False
        
        logger.info(f"Imputed {df_imputed['Is_Imputed'].sum()} values")
        
        return df_imputed
    
    def process_raw_data(self) -> None:
        """
        Complete raw data processing pipeline.
        """
        logger.info("Starting raw data processing pipeline...")
        
        try:
            # Step 1: Parse input data
            logger.info("Processing input data...")
            self.input_df = self.parse_timestamp(self.input_df)
            
            # Step 2: Parse output data
            logger.info("Processing output data...")
            self.output_df = self.parse_timestamp(self.output_df)
            
            # Step 3: Extract structured fields from input
            input_structured = []
            for idx, row in self.input_df.iterrows():
                content = row['Content'] if pd.notna(row['Content']) else ''
                parsed = self.parse_content_structured(content)
                structured_row = row.to_dict()
                structured_row.update(parsed)
                input_structured.append(structured_row)
            
            self.input_df = pd.DataFrame(input_structured)
            
            # Step 4: Extract structured fields from output
            output_structured = []
            for idx, row in self.output_df.iterrows():
                content = row['Content'] if pd.notna(row['Content']) else ''
                parsed = self.parse_content_structured(content)
                structured_row = row.to_dict()
                structured_row.update(parsed)
                output_structured.append(structured_row)
            
            self.output_df = pd.DataFrame(output_structured)
            
            # Step 5: Assign shifts
            for df in [self.input_df, self.output_df]:
                if 'Timestamp' in df.columns:
                    df['Shift'] = df['Timestamp'].apply(self.determine_shift)
                    df['Shift_Date'] = df['Timestamp'].apply(self.assign_shift_date)
                    df['Is_Break_Time'] = df['Timestamp'].apply(self.determine_is_break_time)
            
            # Step 6: Merge data
            self.processed_df = self.merge_input_output_data(self.input_df, self.output_df)
            
            # Step 7: Remove duplicates
            self.processed_df, self.duplicate_df = self.remove_duplicates(self.processed_df)
            
            # Step 8: Validate and remove invalid events
            self.processed_df, self.invalid_df = self.remove_invalid_events(self.processed_df)
            
            # Step 9: Impute missing states
            self.processed_df = self.impute_missing_states(self.processed_df)
            
            # Step 10: Calculate durations
            self.processed_df = self.calculate_duration(self.processed_df)
            
            # Step 11: Remove noise events
            self.processed_df, self.noise_df = self.remove_noise_events(self.processed_df)
            
            # Step 12: Generate calculation audit
            self.generate_calculation_audit()
            
            logger.info("Raw data processing complete!")
            
        except Exception as e:
            logger.error(f"Error in processing pipeline: {str(e)}", exc_info=True)
            raise
    
    def generate_calculation_audit(self) -> None:
        """
        Generate detailed calculation audit trail.
        """
        logger.info("Generating calculation audit...")
        
        audit_records = []
        
        audit_records.append({
            'Step': 'Initial Data Loading',
            'Description': 'Load input and output data files',
            'Value': f"Input: {len(self.input_df) if self.input_df is not None else 0}, Output: {len(self.output_df) if self.output_df is not None else 0}",
            'Units': 'rows',
            'Formula': 'N/A',
            'Source': 'Files loaded',
            'Timestamp': datetime.now().isoformat()
        })
        
        audit_records.append({
            'Step': 'Timestamp Parsing',
            'Description': 'Parse timestamps from string to datetime',
            'Value': f"{len(self.input_df) + len(self.output_df) if self.input_df is not None and self.output_df is not None else 0} timestamps parsed",
            'Units': 'events',
            'Formula': 'pd.to_datetime(Created_At)',
            'Source': 'Created_At column',
            'Timestamp': datetime.now().isoformat()
        })
        
        audit_records.append({
            'Step': 'Machine ID Extraction',
            'Description': 'Extract machine ID from content',
            'Value': 'Machine 1 and Machine 2 identified',
            'Units': 'machines',
            'Formula': 'regex extraction',
            'Source': 'Content column',
            'Timestamp': datetime.now().isoformat()
        })
        
        if self.processed_df is not None:
            person_count_events = len(self.processed_df[self.processed_df['Event_Type'] == 'Personnel_Count'])
            audit_records.append({
                'Step': 'Person Count Extraction',
                'Description': 'Extract person counts from content',
                'Value': f"{person_count_events} personnel events",
                'Units': 'events',
                'Formula': 'regex extraction',
                'Source': 'Content column',
                'Timestamp': datetime.now().isoformat()
            })
            
            shift_counts = self.processed_df['Shift'].value_counts().to_dict()
            audit_records.append({
                'Step': 'Shift Assignment',
                'Description': 'Assign events to shifts with break handling',
                'Value': f"Shift 1: {shift_counts.get('Shift 1', 0)}, Shift 2: {shift_counts.get('Shift 2', 0)}, Shift 3: {shift_counts.get('Shift 3', 0)}",
                'Units': 'events',
                'Formula': 'determine_shift(timestamp)',
                'Source': 'Timestamp column',
                'Timestamp': datetime.now().isoformat()
            })
            
            if self.duplicate_df is not None:
                audit_records.append({
                    'Step': 'Duplicate Removal',
                    'Description': 'Remove duplicate events based on timestamp and machine',
                    'Value': f"{len(self.duplicate_df)} duplicates removed",
                    'Units': 'events',
                    'Formula': 'drop_duplicates(subset=[timestamp, machine_id])',
                    'Source': 'Processed data',
                    'Timestamp': datetime.now().isoformat()
                })
            
            if self.invalid_df is not None:
                audit_records.append({
                    'Step': 'Invalid Event Removal',
                    'Description': 'Remove invalid events based on validation rules',
                    'Value': f"{len(self.invalid_df)} invalid events removed",
                    'Units': 'events',
                    'Formula': 'validate_event(row)',
                    'Source': 'Processed data',
                    'Timestamp': datetime.now().isoformat()
                })
            
            if self.noise_df is not None:
                audit_records.append({
                    'Step': 'Noise Event Removal',
                    'Description': f'Remove short-duration noise events (< {self.config.min_event_duration_seconds}s)',
                    'Value': f"{len(self.noise_df)} noise events removed",
                    'Units': 'events',
                    'Formula': f'duration < {self.config.min_event_duration_seconds}s',
                    'Source': 'Adjusted_Duration_Seconds column',
                    'Timestamp': datetime.now().isoformat()
                })
            
            total_duration = self.processed_df['Duration_Seconds'].sum() / 3600 if len(self.processed_df) > 0 else 0
            total_adjusted = self.processed_df['Adjusted_Duration_Seconds'].sum() / 3600 if len(self.processed_df) > 0 else 0
            audit_records.append({
                'Step': 'Duration Calculation',
                'Description': 'Calculate event durations and adjust for breaks',
                'Value': f"Total: {total_duration:.2f}h, Adjusted: {total_adjusted:.2f}h",
                'Units': 'hours',
                'Formula': 'next_timestamp - current_timestamp',
                'Source': 'Timestamp column',
                'Timestamp': datetime.now().isoformat()
            })
        
        self.calculation_audit = pd.DataFrame(audit_records)
        logger.info(f"Calculation audit generated: {len(self.calculation_audit)} records")
    
    def generate_processed_data_sheet(self) -> pd.DataFrame:
        """
        Generate processed data sheet for export.
        
        Returns:
            DataFrame with processed data in export format
        """
        logger.info("Generating processed data sheet...")
        
        if self.processed_df is None or len(self.processed_df) == 0:
            logger.warning("No processed data available")
            return pd.DataFrame()
        
        export_columns = [
            'Timestamp',
            'Date',
            'Time',
            'Shift',
            'Shift_Date',
            'Is_Break_Time',
            'Machine_ID',
            'Side',
            'Source',
            'Event_Type',
            'Machine_Status',
            'Machine_Status_Imputed',
            'Person_Count',
            'Person_Count_Imputed',
            'Has_Box',
            'Has_Box_Imputed',
            'Has_No_Box',
            'Has_No_Box_Imputed',
            'Is_Feeding',
            'Is_Bringing_Material',
            'Operator_Present',
            'Operator_Not_Present',
            'Material_Present_Operator_Absent',
            'Operator_Present_Material_Absent',
            'Process_Statuses',
            'Content',
            'Video_Link',
            'Duration_Seconds',
            'Duration_Minutes',
            'Break_Overlap_Seconds',
            'Adjusted_Duration_Seconds',
            'Adjusted_Duration_Minutes',
            'Is_Imputed',
            'Is_Valid',
            'Invalid_Reason'
        ]
        
        existing_columns = [col for col in export_columns if col in self.processed_df.columns]
        processed_export = self.processed_df[existing_columns].copy()
        
        if 'Timestamp' in processed_export.columns:
            processed_export['Timestamp'] = processed_export['Timestamp'].dt.strftime('%Y-%m-%d %H:%M:%S')
        if 'Date' in processed_export.columns:
            processed_export['Date'] = processed_export['Date'].astype(str)
        if 'Time' in processed_export.columns:
            processed_export['Time'] = processed_export['Time'].astype(str)
        if 'Shift_Date' in processed_export.columns:
            processed_export['Shift_Date'] = processed_export['Shift_Date'].astype(str)
        
        if 'Process_Statuses' in processed_export.columns:
            processed_export['Process_Statuses'] = processed_export['Process_Statuses'].apply(
                lambda x: ', '.join(x) if isinstance(x, list) else str(x) if pd.notna(x) else ''
            )
        
        logger.info(f"Processed data sheet: {len(processed_export)} rows")
        
        return processed_export
    
    def generate_raw_data_sheet(self) -> pd.DataFrame:
        """
        Generate raw data sheet for export.
        
        Returns:
            DataFrame with raw data
        """
        logger.info("Generating raw data sheet...")
        
        if self.processed_df is None or len(self.processed_df) == 0:
            logger.warning("No processed data available")
            return pd.DataFrame()
        
        raw_columns = []
        
        if 'Timestamp' in self.processed_df.columns:
            raw_columns.append('Timestamp')
        if 'Created_At' in self.processed_df.columns:
            raw_columns.append('Created_At')
        if 'Content' in self.processed_df.columns:
            raw_columns.append('Content')
        if 'Video_Link' in self.processed_df.columns:
            raw_columns.append('Video_Link')
        if 'Machine_ID' in self.processed_df.columns:
            raw_columns.append('Machine_ID')
        if 'Side' in self.processed_df.columns:
            raw_columns.append('Side')
        if 'Source' in self.processed_df.columns:
            raw_columns.append('Source')
        if 'Shift' in self.processed_df.columns:
            raw_columns.append('Shift')
        if 'Event_Type' in self.processed_df.columns:
            raw_columns.append('Event_Type')
        
        raw_export = self.processed_df[raw_columns].copy()
        
        if 'Timestamp' in raw_export.columns:
            raw_export['Timestamp'] = raw_export['Timestamp'].dt.strftime('%Y-%m-%d %H:%M:%S')
        if 'Created_At' in raw_export.columns:
            raw_export['Created_At'] = pd.to_datetime(raw_export['Created_At'], errors='coerce').dt.strftime('%Y-%m-%d %H:%M:%S')
        
        logger.info(f"Raw data sheet: {len(raw_export)} rows")
        
        return raw_export
    
    def generate_invalid_events_sheet(self) -> pd.DataFrame:
        """
        Generate invalid events sheet for export.
        
        Returns:
            DataFrame with invalid events
        """
        logger.info("Generating invalid events sheet...")
        
        if self.invalid_df is None or len(self.invalid_df) == 0:
            logger.info("No invalid events found")
            return pd.DataFrame()
        
        invalid_export = self.invalid_df.copy()
        
        if 'Timestamp' in invalid_export.columns:
            invalid_export['Timestamp'] = invalid_export['Timestamp'].dt.strftime('%Y-%m-%d %H:%M:%S')
        
        export_columns = []
        if 'Timestamp' in invalid_export.columns:
            export_columns.append('Timestamp')
        if 'Content' in invalid_export.columns:
            export_columns.append('Content')
        if 'Video_Link' in invalid_export.columns:
            export_columns.append('Video_Link')
        if 'Machine_ID' in invalid_export.columns:
            export_columns.append('Machine_ID')
        if 'Side' in invalid_export.columns:
            export_columns.append('Side')
        if 'Event_Type' in invalid_export.columns:
            export_columns.append('Event_Type')
        if 'Invalid_Reason' in invalid_export.columns:
            export_columns.append('Invalid_Reason')
        
        invalid_export = invalid_export[export_columns]
        
        logger.info(f"Invalid events sheet: {len(invalid_export)} rows")
        
        return invalid_export
    
    def generate_noise_events_sheet(self) -> pd.DataFrame:
        """
        Generate noise events sheet for export.
        
        Returns:
            DataFrame with noise events
        """
        logger.info("Generating noise events sheet...")
        
        if self.noise_df is None or len(self.noise_df) == 0:
            logger.info("No noise events found")
            return pd.DataFrame()
        
        noise_export = self.noise_df.copy()
        
        if 'Timestamp' in noise_export.columns:
            noise_export['Timestamp'] = noise_export['Timestamp'].dt.strftime('%Y-%m-%d %H:%M:%S')
        
        export_columns = []
        if 'Timestamp' in noise_export.columns:
            export_columns.append('Timestamp')
        if 'Content' in noise_export.columns:
            export_columns.append('Content')
        if 'Machine_ID' in noise_export.columns:
            export_columns.append('Machine_ID')
        if 'Side' in noise_export.columns:
            export_columns.append('Side')
        if 'Adjusted_Duration_Seconds' in noise_export.columns:
            export_columns.append('Adjusted_Duration_Seconds')
        if 'Event_Type' in noise_export.columns:
            export_columns.append('Event_Type')
        
        noise_export = noise_export[export_columns]
        
        logger.info(f"Noise events sheet: {len(noise_export)} rows")
        
        return noise_export
    
    def generate_config_sheet(self) -> pd.DataFrame:
        """
        Generate configuration sheet for export.
        
        Returns:
            DataFrame with configuration parameters
        """
        logger.info("Generating configuration sheet...")
        
        config_data = []
        
        for shift_name, shift_data in self.config.shifts.items():
            config_data.append({
                'Parameter': f'{shift_name} Start',
                'Value': shift_data['start'].strftime('%H:%M'),
                'Description': f'{shift_name} start time'
            })
            config_data.append({
                'Parameter': f'{shift_name} End',
                'Value': shift_data['end'].strftime('%H:%M'),
                'Description': f'{shift_name} end time'
            })
            config_data.append({
                'Parameter': f'{shift_name} Break Start',
                'Value': shift_data['break_start'].strftime('%H:%M'),
                'Description': f'{shift_name} break start time'
            })
            config_data.append({
                'Parameter': f'{shift_name} Break End',
                'Value': shift_data['break_end'].strftime('%H:%M'),
                'Description': f'{shift_name} break end time'
            })
        
        config_data.append({
            'Parameter': 'Required Input Operators',
            'Value': self.config.required_input_operators,
            'Description': 'Number of operators required on input side'
        })
        config_data.append({
            'Parameter': 'Required Output Operators',
            'Value': self.config.required_output_operators,
            'Description': 'Number of operators required on output side'
        })
        
        config_data.append({
            'Parameter': 'Min Event Duration (seconds)',
            'Value': self.config.min_event_duration_seconds,
            'Description': 'Minimum event duration for noise filtering'
        })
        config_data.append({
            'Parameter': 'Max Imputation Gap (seconds)',
            'Value': self.config.max_gap_for_imputation_seconds,
            'Description': 'Maximum gap for imputing missing states'
        })
        config_data.append({
            'Parameter': 'Duplicate Tolerance (seconds)',
            'Value': self.config.duplicate_tolerance_seconds,
            'Description': 'Time tolerance for duplicate detection'
        })
        
        config_df = pd.DataFrame(config_data)
        logger.info(f"Configuration sheet: {len(config_df)} parameters")
        
        return config_df
    
    def generate_calculation_audit_sheet(self) -> pd.DataFrame:
        """
        Generate calculation audit sheet for export.
        
        Returns:
            DataFrame with calculation audit records
        """
        logger.info("Generating calculation audit sheet...")
        
        if self.calculation_audit is None:
            return pd.DataFrame()
        
        return self.calculation_audit
    
    def save_to_excel(self, output_file: str = 'processed_data_script1.xlsx') -> None:
        """
        Save all processed data to Excel workbook.
        
        Args:
            output_file: Output Excel file path
        """
        logger.info(f"Saving data to Excel: {output_file}")
        
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Sheet 1: Config
                config_df = self.generate_config_sheet()
                if not config_df.empty:
                    config_df.to_excel(writer, sheet_name='Config', index=False)
                
                # Sheet 2: Raw Data
                raw_df = self.generate_raw_data_sheet()
                if not raw_df.empty:
                    raw_df.to_excel(writer, sheet_name='Raw Data', index=False)
                
                # Sheet 3: Processed Data
                processed_df = self.generate_processed_data_sheet()
                if not processed_df.empty:
                    processed_df.to_excel(writer, sheet_name='Processed Data', index=False)
                
                # Sheet 4: Invalid Events
                invalid_df = self.generate_invalid_events_sheet()
                if not invalid_df.empty:
                    invalid_df.to_excel(writer, sheet_name='Invalid Events', index=False)
                
                # Sheet 5: Noise Events
                noise_df = self.generate_noise_events_sheet()
                if not noise_df.empty:
                    noise_df.to_excel(writer, sheet_name='Noise Events', index=False)
                
                # Sheet 6: Calculation Audit
                audit_df = self.generate_calculation_audit_sheet()
                if not audit_df.empty:
                    audit_df.to_excel(writer, sheet_name='Calculation Audit', index=False)
                
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


def main():
    """
    Main execution function.
    """
    logger.info("=" * 80)
    logger.info("DATA PROCESSING SCRIPT - SCRIPT 1 OF PROJECT")
    logger.info("=" * 80)
    
    try:
        config = DataProcessingConfig()
        logger.info("Configuration initialized successfully")
        
        processor = DataProcessor(config)
        logger.info("Data processor initialized")
        
        input_file = 'Machines Input Side Data 22 May 2026 to 23 May 2026.xlsx'
        output_file = 'Machines Output Side Data 22 May 2026 to 23 May 2026.xlsx'
        
        input_path = Path(input_file)
        output_path = Path(output_file)
        
        if not input_path.exists():
            logger.error(f"Input file not found: {input_file}")
            logger.info("Please ensure the file is in the current directory")
            return
        
        if not output_path.exists():
            logger.error(f"Output file not found: {output_file}")
            logger.info("Please ensure the file is in the current directory")
            return
        
        processor.load_data(input_file, output_file)
        processor.process_raw_data()
        
        output_excel = 'processed_data_script1_final.xlsx'
        processor.save_to_excel(output_excel)
        
        logger.info("=" * 80)
        logger.info("SCRIPT 1 EXECUTION COMPLETED SUCCESSFULLY")
        logger.info(f"OUTPUT FILE: {output_excel}")
        logger.info("=" * 80)
        
        logger.info("\nSUMMARY STATISTICS:")
        if processor.processed_df is not None:
            logger.info(f"Total events processed: {len(processor.processed_df)}")
        else:
            logger.info("Total events processed: 0")
        
        if processor.input_df is not None:
            logger.info(f"Input events: {len(processor.input_df)}")
        else:
            logger.info("Input events: 0")
        
        if processor.output_df is not None:
            logger.info(f"Output events: {len(processor.output_df)}")
        else:
            logger.info("Output events: 0")
        
        if processor.duplicate_df is not None:
            logger.info(f"Duplicates removed: {len(processor.duplicate_df)}")
        else:
            logger.info("Duplicates removed: 0")
        
        if processor.invalid_df is not None:
            logger.info(f"Invalid events: {len(processor.invalid_df)}")
        else:
            logger.info("Invalid events: 0")
        
        if processor.noise_df is not None:
            logger.info(f"Noise events removed: {len(processor.noise_df)}")
        else:
            logger.info("Noise events removed: 0")
        
        if processor.processed_df is not None and 'Shift' in processor.processed_df.columns:
            logger.info("Shift distribution:")
            shift_counts = processor.processed_df['Shift'].value_counts()
            for shift, count in shift_counts.items():
                logger.info(f"  {shift}: {count} events")
        
        if processor.processed_df is not None and 'Duration_Seconds' in processor.processed_df.columns:
            total_duration = processor.processed_df['Duration_Seconds'].sum() / 3600
            logger.info(f"Total duration: {total_duration:.2f} hours")
            
            if 'Adjusted_Duration_Seconds' in processor.processed_df.columns:
                total_adjusted = processor.processed_df['Adjusted_Duration_Seconds'].sum() / 3600
                logger.info(f"Adjusted duration (excluding breaks): {total_adjusted:.2f} hours")
        
    except Exception as e:
        logger.error(f"Critical error in main execution: {str(e)}", exc_info=True)
        raise


if __name__ == "__main__":
    main()