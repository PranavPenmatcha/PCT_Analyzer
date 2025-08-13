#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pulse Analysis Tool for WinDaq Current Data

This script analyzes current pulses in the Excel file, detects peaks,
and adds analysis results to a new sheet in the Excel file.
"""

import pandas as pd
import numpy as np
from pathlib import Path
import logging
from scipy.signal import find_peaks, peak_widths
import matplotlib.pyplot as plt

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class PulseAnalyzer:
    def __init__(self, excel_file_path):
        """
        Initialize the pulse analyzer with Excel data.
        
        Args:
            excel_file_path (str): Path to the Excel file containing time and current data
        """
        self.excel_file_path = excel_file_path
        self.df = None
        self.pulses = []
        
    def load_data(self):
        """Load data from Excel file."""
        logger.info(f"Loading data from: {self.excel_file_path}")
        self.df = pd.read_excel(self.excel_file_path)
        
        # Verify required columns exist
        if 'Time_seconds' not in self.df.columns or 'Current' not in self.df.columns:
            raise ValueError("Excel file must contain 'Time_seconds' and 'Current' columns")
        
        logger.info(f"Loaded {len(self.df)} data points")
        logger.info(f"Time range: {self.df['Time_seconds'].min():.3f} to {self.df['Time_seconds'].max():.3f} seconds")
        logger.info(f"Current range: {self.df['Current'].min():.2f} to {self.df['Current'].max():.2f} A")
        
    def detect_pulses(self, threshold_percent=10, min_pulse_width=0.001, min_pulse_separation=0.01):
        """
        Detect current pulses in the data.
        
        Args:
            threshold_percent (float): Minimum current as percentage of max current to consider a pulse
            min_pulse_width (float): Minimum pulse width in seconds
            min_pulse_separation (float): Minimum separation between pulses in seconds
        
        Returns:
            list: List of pulse dictionaries with start, end, peak info
        """
        logger.info("Detecting current pulses...")
        
        current = self.df['Current'].values
        time = self.df['Time_seconds'].values
        
        # Calculate threshold
        max_current = np.max(current)
        threshold = max_current * (threshold_percent / 100)
        logger.info(f"Using threshold: {threshold:.2f} A ({threshold_percent}% of max current {max_current:.2f} A)")
        
        # Find points above threshold
        above_threshold = current > threshold
        
        # Find pulse boundaries
        pulse_starts = []
        pulse_ends = []
        
        in_pulse = False
        for i, above in enumerate(above_threshold):
            if above and not in_pulse:
                # Start of pulse
                pulse_starts.append(i)
                in_pulse = True
            elif not above and in_pulse:
                # End of pulse
                pulse_ends.append(i-1)
                in_pulse = False
        
        # Handle case where data ends during a pulse
        if in_pulse:
            pulse_ends.append(len(current) - 1)
        
        # Filter pulses by minimum width and separation
        sample_rate = 1 / (time[1] - time[0])
        min_width_samples = int(min_pulse_width * sample_rate)
        min_separation_samples = int(min_pulse_separation * sample_rate)
        
        filtered_pulses = []
        for start_idx, end_idx in zip(pulse_starts, pulse_ends):
            pulse_width = end_idx - start_idx
            
            # Check minimum width
            if pulse_width < min_width_samples:
                continue
            
            # Check separation from previous pulse
            if filtered_pulses:
                prev_end = filtered_pulses[-1]['end_idx']
                if start_idx - prev_end < min_separation_samples:
                    continue
            
            # Find peak within pulse
            pulse_current = current[start_idx:end_idx+1]
            peak_idx_relative = np.argmax(pulse_current)
            peak_idx = start_idx + peak_idx_relative
            
            pulse_info = {
                'pulse_number': len(filtered_pulses) + 1,
                'start_idx': start_idx,
                'end_idx': end_idx,
                'peak_idx': peak_idx,
                'start_time': time[start_idx],
                'end_time': time[end_idx],
                'peak_time': time[peak_idx],
                'duration': time[end_idx] - time[start_idx],
                'peak_current': current[peak_idx],
                'avg_current': np.mean(pulse_current),
                'pulse_energy': np.trapz(pulse_current, time[start_idx:end_idx+1])
            }
            
            filtered_pulses.append(pulse_info)
        
        self.pulses = filtered_pulses
        logger.info(f"Detected {len(self.pulses)} pulses")
        
        return self.pulses
    
    def analyze_pulse_statistics(self):
        """
        Analyze statistics of detected pulses.
        
        Returns:
            dict: Dictionary containing pulse statistics
        """
        if not self.pulses:
            logger.warning("No pulses detected. Run detect_pulses() first.")
            return {}
        
        peak_currents = [p['peak_current'] for p in self.pulses]
        durations = [p['duration'] for p in self.pulses]
        energies = [p['pulse_energy'] for p in self.pulses]
        
        stats = {
            'total_pulses': len(self.pulses),
            'peak_current_min': np.min(peak_currents),
            'peak_current_max': np.max(peak_currents),
            'peak_current_mean': np.mean(peak_currents),
            'peak_current_std': np.std(peak_currents),
            'duration_min': np.min(durations),
            'duration_max': np.max(durations),
            'duration_mean': np.mean(durations),
            'duration_std': np.std(durations),
            'energy_total': np.sum(energies),
            'energy_mean': np.mean(energies),
            'first_pulse_time': self.pulses[0]['start_time'],
            'last_pulse_time': self.pulses[-1]['end_time'],
            'test_duration': self.pulses[-1]['end_time'] - self.pulses[0]['start_time']
        }
        
        logger.info(f"Pulse Statistics:")
        logger.info(f"  Total pulses: {stats['total_pulses']}")
        logger.info(f"  Peak current range: {stats['peak_current_min']:.2f} - {stats['peak_current_max']:.2f} A")
        logger.info(f"  Average peak current: {stats['peak_current_mean']:.2f} ± {stats['peak_current_std']:.2f} A")
        logger.info(f"  Pulse duration range: {stats['duration_min']:.4f} - {stats['duration_max']:.4f} s")
        logger.info(f"  Average pulse duration: {stats['duration_mean']:.4f} ± {stats['duration_std']:.4f} s")
        
        return stats
    
    def create_pulse_summary_sheet(self):
        """
        Create a summary sheet with pulse analysis results.
        
        Returns:
            pd.DataFrame: DataFrame containing pulse summary
        """
        if not self.pulses:
            logger.warning("No pulses detected. Run detect_pulses() first.")
            return pd.DataFrame()
        
        # Create pulse summary DataFrame
        pulse_data = []
        for pulse in self.pulses:
            pulse_data.append({
                'Pulse_Number': pulse['pulse_number'],
                'Start_Time_s': round(pulse['start_time'], 4),
                'Peak_Time_s': round(pulse['peak_time'], 4),
                'End_Time_s': round(pulse['end_time'], 4),
                'Duration_s': round(pulse['duration'], 4),
                'Peak_Current_A': round(pulse['peak_current'], 2),
                'Average_Current_A': round(pulse['avg_current'], 2),
                'Pulse_Energy_J': round(pulse['pulse_energy'], 4)
            })
        
        pulse_df = pd.DataFrame(pulse_data)
        
        # Add statistics summary
        stats = self.analyze_pulse_statistics()
        
        # Create statistics DataFrame
        stats_data = [
            ['Total Pulses', stats['total_pulses'], ''],
            ['Peak Current Min (A)', round(stats['peak_current_min'], 2), ''],
            ['Peak Current Max (A)', round(stats['peak_current_max'], 2), ''],
            ['Peak Current Mean (A)', round(stats['peak_current_mean'], 2), ''],
            ['Peak Current Std (A)', round(stats['peak_current_std'], 2), ''],
            ['Duration Min (s)', round(stats['duration_min'], 4), ''],
            ['Duration Max (s)', round(stats['duration_max'], 4), ''],
            ['Duration Mean (s)', round(stats['duration_mean'], 4), ''],
            ['Duration Std (s)', round(stats['duration_std'], 4), ''],
            ['Total Energy (J)', round(stats['energy_total'], 2), ''],
            ['Average Energy per Pulse (J)', round(stats['energy_mean'], 4), ''],
            ['Test Duration (s)', round(stats['test_duration'], 4), ''],
        ]
        
        stats_df = pd.DataFrame(stats_data, columns=['Statistic', 'Value', 'Unit'])
        
        return pulse_df, stats_df
    
    def save_analysis_to_excel(self, output_file_path=None):
        """
        Save the analysis results to Excel file with pulse summary in Raw_Data sheet.

        Args:
            output_file_path (str, optional): Output file path. If None, modifies original file.
        """
        if output_file_path is None:
            output_file_path = self.excel_file_path

        logger.info(f"Saving analysis results to: {output_file_path}")

        if not self.pulses:
            logger.warning("No pulses detected. Saving original data only.")
            self.df.to_excel(output_file_path, index=False)
            return output_file_path

        # Create a simple pulse summary table
        pulse_summary_data = []
        pulse_summary_data.append(['Pulse', 'Peak Current'])  # Header row

        for pulse in self.pulses:
            pulse_summary_data.append([
                f'Pulse {pulse["pulse_number"]}',
                round(pulse['peak_current'], 0)  # Round to nearest whole number
            ])

        # Create DataFrame for the summary table
        summary_df = pd.DataFrame(pulse_summary_data, columns=['Pulse', 'Peak Current'])

        # Write to Excel with the summary table added to the right of the raw data
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            # Write original data
            self.df.to_excel(writer, sheet_name='Raw_Data', index=False, startcol=0)

            # Write pulse summary table starting from column D (index 3)
            summary_df.to_excel(writer, sheet_name='Raw_Data', index=False, startcol=3)

        logger.info(f"Analysis complete! Results saved to: {output_file_path}")
        logger.info("Pulse summary table added to the right of the raw data in Raw_Data sheet")
        return output_file_path


def main():
    """Main function to run pulse analysis."""
    # Find Excel file
    excel_files = list(Path('.').glob('*.xlsx'))
    
    if not excel_files:
        logger.error("No Excel files found in current directory")
        return
    
    excel_file = excel_files[0]
    logger.info(f"Analyzing file: {excel_file}")
    
    try:
        # Create analyzer
        analyzer = PulseAnalyzer(str(excel_file))
        
        # Load data
        analyzer.load_data()
        
        # Detect pulses
        pulses = analyzer.detect_pulses(threshold_percent=5, min_pulse_width=0.001, min_pulse_separation=0.01)
        
        if pulses:
            # Analyze statistics
            stats = analyzer.analyze_pulse_statistics()
            
            # Save results
            output_file = analyzer.save_analysis_to_excel()
            
            print(f"\nPulse Analysis Complete!")
            print(f"Found {len(pulses)} current pulses")
            print(f"Peak current range: {stats['peak_current_min']:.2f} - {stats['peak_current_max']:.2f} A")
            print(f"Highest peak: {stats['peak_current_max']:.2f} A")
            print(f"Results saved to: {output_file}")
            print(f"Pulse summary table added to the right of the raw data")

            # Show the pulse summary
            print(f"\nPulse Summary:")
            print(f"              Peak Current")
            for pulse in pulses:
                print(f"Pulse {pulse['pulse_number']}      {pulse['peak_current']:.0f}")

            # Show detailed pulse information for web parsing
            print(f"\nIndividual Pulse Details:")
            for pulse in pulses:
                print(f"Pulse {pulse['pulse_number']}: {pulse['peak_current']:.2f} A at {pulse['peak_time']:.3f}s")
            
        else:
            print("No pulses detected. Try adjusting the threshold parameters.")
            
    except Exception as e:
        logger.error(f"Error during analysis: {str(e)}")
        raise


if __name__ == "__main__":
    main()
