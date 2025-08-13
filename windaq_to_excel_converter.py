#!/usr/bin/env python3
"""
WinDaq to Excel Converter

This script converts WinDaq files (.wdq, .wdh, .wdc) to Excel format.
It supports two methods:
1. Converting CSV files exported from WinDaq Waveform Browser
2. Direct reading of WinDaq files using Python libraries (experimental)

Usage:
    python windaq_to_excel_converter.py <input_file> [output_file]

Examples:
    python windaq_to_excel_converter.py data.csv
    python windaq_to_excel_converter.py data.wdq output.xlsx
"""

import os
import sys
import argparse
import pandas as pd
from pathlib import Path
import logging
import struct
import datetime
import numpy

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class windaq(object):
    '''
    Read windaq files (.wdq extension) without having to convert them to .csv or other human readable text
    Code based on http://www.dataq.com/resources/pdfs/misc/ff.pdf provided by Dataq
    '''

    def __init__(self, filename):
        '''
        Define data types based off convention used in documentation from Dataq
        '''
        UI = "<H"  # unsigned int 16 bit
        UL = "<L"  # unsigned long 32 bit
        I = "<h"   # signed int 16 bit
        L = "<l"   # signed long 32 bit
        F = "<f"   # float 32 bit
        D = "<d"   # double 64 bit
        B = "<B"   # unsigned char 8 bit

        '''
        Read file into memory
        '''
        with open(filename, 'rb') as f:
            self._fcontents = f.read()

        '''
        Read header information
        '''
        # element 0 contains packed info about number of channels and file type
        if struct.unpack_from(UI, self._fcontents, 2)[0] >= 144:
            self.nChannels = (struct.unpack_from(B, self._fcontents, 0)[0])  # number of channels is element 1
        else:
            self.nChannels = (struct.unpack_from(B, self._fcontents, 0)[0]) & 31  # number of channels is element 1 mask bit 5

        self._hChannels = struct.unpack_from(B, self._fcontents, 4)[0]  # offset in bytes from BOF to header channel info tables
        self._hChannelSize = struct.unpack_from(B, self._fcontents, 5)[0]  # number of bytes in each channel info entry
        self._headSize = struct.unpack_from(I, self._fcontents, 6)[0]  # number of bytes in data file header
        self._dataSize = struct.unpack_from(UL, self._fcontents, 8)[0]  # number of ADC data bytes in file excluding header
        self.nSample = (self._dataSize/(2*self.nChannels))  # number of samples per channel
        self._trailerSize = struct.unpack_from(UL, self._fcontents, 12)[0]  # total number of event marker, time and date stamp, and event marker comment pointer bytes in trailer
        self._annoSize = struct.unpack_from(UI, self._fcontents, 16)[0]  # total number of usr annotation bytes including 1 null per channel
        self.timeStep = struct.unpack_from(D, self._fcontents, 28)[0]  # time between channel samples: 1/(sample rate throughput / total number of acquired channels)

        e14 = struct.unpack_from(L, self._fcontents, 36)[0]  # time file was opened by acquisition: total number of seconds since jan 1 1970
        e15 = struct.unpack_from(L, self._fcontents, 40)[0]  # time file was written by acquisition: total number of seconds since jan 1 1970

        self.fileCreatedRaw = datetime.datetime.utcfromtimestamp(e14)  # datetime format of time file was opened by acquisition
        self.fileCreated = datetime.datetime.fromtimestamp(e14).strftime('%Y-%m-%d %H:%M:%S')  # datetime format of time file was opened by acquisition
        self.fileWritten = datetime.datetime.fromtimestamp(e15).strftime('%Y-%m-%d %H:%M:%S')  # datetime format of time file was written by acquisition

        self._packed = ((struct.unpack_from(UI, self._fcontents, 100)[0]) & 16384) >> 14  # bit 14 of element 27 indicates packed file. bitwise & e27 with 16384 to mask all bits but 14 and then shift to 0 bit place
        self._HiRes = ((struct.unpack_from(UI, self._fcontents, 100)[0]) & 2)  # bit 1 of element 27 indicates a HiRes file with 16-bit data

        '''
        read channel info
        '''
        self.scalingSlope = []
        self.scalingIntercept = []
        self.calScaling = []
        self.calIntercept = []
        self.engUnits = []
        self.sampleRateDivisor = []
        self.phyChannel = []

        for channel in range(0, self.nChannels):
            channelOffset = self._hChannels + (self._hChannelSize * channel)  # calculate channel header offset from beginning of file, each channel header size is defined in _hChannelSize
            self.scalingSlope.append(struct.unpack_from(F, self._fcontents, channelOffset)[0])  # scaling slope (m) applied to the waveform to scale it within the display window
            self.scalingIntercept.append(struct.unpack_from(F, self._fcontents, channelOffset + 4)[0])  # scaling intercept (b) applied to the waveform to scale it withing the display window
            self.calScaling.append(struct.unpack_from(D, self._fcontents, channelOffset + 4 + 4)[0])  # calibration scaling factor (m) for waveform value display
            self.calIntercept.append(struct.unpack_from(D, self._fcontents, channelOffset + 4 + 4 + 8)[0])  # calibration intercept factor (b) for waveform value display
            self.engUnits.append(struct.unpack_from("cccccc", self._fcontents, channelOffset + 4 + 4 + 8 + 8))  # engineering units tag for calibrated waveform, only 4 bits are used last two are null

            if self._packed:  # if file is packed then item 7 is the sample rate divisor
                self.sampleRateDivisor.append(struct.unpack_from(B, self._fcontents, channelOffset + 4 + 4 + 8 + 8 + 6 + 1)[0])
            else:
                self.sampleRateDivisor.append(1)

            self.phyChannel.append(struct.unpack_from(B, self._fcontents, channelOffset + 4 + 4 + 8 + 8 + 6 + 1 + 1)[0])  # describes the physical channel number

        '''
        read user annotations
        '''
        aOffset = self._headSize + self._dataSize + self._trailerSize
        aTemp = ''
        for i in range(0, self._annoSize):
            aTemp += struct.unpack_from('c', self._fcontents, aOffset + i)[0].decode("utf-8")
        self._annotations = aTemp.split('\x00')

        # create a numpy view into the data for efficient reading
        dt = numpy.dtype(numpy.int16)
        dt = dt.newbyteorder('<')
        self.npdata = numpy.frombuffer(self._fcontents, dtype=dt, count=int(self.nSample*self.nChannels), offset=self._headSize)

    def data(self, channelNumber):
        '''
        return the data for the channel requested
        data format is saved CH1tonChannels one sample at a time. each sample is read as a 16bit word and then shifted to a 14bit value
        '''
        data = self.npdata[(channelNumber-1)::self.nChannels]

        if self._HiRes:
            temp = data * 0.25  # multiply by 0.25 for HiRes data
        else:
            temp = numpy.floor(data*0.25)  # bit shift by two for normal data

        temp2 = self.calScaling[channelNumber-1]*temp + self.calIntercept[channelNumber-1]
        return temp2

    def time(self):
        '''
        return time (relative to logger start time)
        '''
        return numpy.arange(0, int(self.nSample))*self.timeStep

    def time_utc(self):
        '''
        return time in numpy datetime64 format
        '''
        return (self.time()*1e9).astype('timedelta64[ns]') + numpy.datetime64(self.fileCreatedRaw)

    def unit(self, channelNumber):
        '''
        return unit of requested channel
        '''
        unit = ''
        for b in self.engUnits[channelNumber-1]:
            unit += b.decode('utf-8')

        '''
        Was getting \x00 in the unit string after decoding, lets remove that and whitespace
        '''
        unit = unit.replace('\x00', '').strip()
        return unit

    def chAnnotation(self, channelNumber):
        '''
        return user annotation of requested channel
        '''
        return self._annotations[channelNumber-1]


def convert_csv_to_excel(csv_file_path, excel_file_path=None):
    """
    Convert a CSV file (exported from WinDaq Waveform Browser) to Excel format.
    
    Args:
        csv_file_path (str): Path to the input CSV file
        excel_file_path (str, optional): Path for the output Excel file
    
    Returns:
        str: Path to the created Excel file
    """
    try:
        # Generate output filename if not provided
        if excel_file_path is None:
            csv_path = Path(csv_file_path)
            excel_file_path = csv_path.with_suffix('.xlsx')
        
        logger.info(f"Reading CSV file: {csv_file_path}")
        
        # Read the CSV file with various encoding attempts
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        df = None
        
        for encoding in encodings:
            try:
                df = pd.read_csv(csv_file_path, encoding=encoding)
                logger.info(f"Successfully read CSV with {encoding} encoding")
                break
            except UnicodeDecodeError:
                continue
        
        if df is None:
            raise ValueError("Could not read CSV file with any supported encoding")
        
        logger.info(f"CSV file contains {len(df)} rows and {len(df.columns)} columns")
        logger.info(f"Columns: {list(df.columns)}")
        
        # Save to Excel
        logger.info(f"Saving to Excel file: {excel_file_path}")
        df.to_excel(excel_file_path, index=False, engine='openpyxl')
        
        logger.info(f"Successfully converted '{csv_file_path}' to '{excel_file_path}'")
        return str(excel_file_path)
        
    except Exception as e:
        logger.error(f"Error converting CSV to Excel: {str(e)}")
        raise


def convert_windaq_direct(wdq_file_path, excel_file_path=None):
    """
    Convert WinDaq files directly using the proper WinDaq parser.
    Creates Excel file with Time and Current columns.

    Args:
        wdq_file_path (str): Path to the input WinDaq file
        excel_file_path (str, optional): Path for the output Excel file

    Returns:
        str: Path to the created Excel file
    """
    try:
        # Generate output filename if not provided
        if excel_file_path is None:
            wdq_path = Path(wdq_file_path)
            excel_file_path = wdq_path.with_suffix('.xlsx')

        logger.info(f"Reading WinDaq file: {wdq_file_path}")

        # Use the proper WinDaq parser
        wdq = windaq(wdq_file_path)

        logger.info(f"File info:")
        logger.info(f"  Number of channels: {wdq.nChannels}")
        logger.info(f"  Number of samples per channel: {int(wdq.nSample)}")
        logger.info(f"  Time step: {wdq.timeStep} seconds")
        logger.info(f"  File created: {wdq.fileCreated}")
        logger.info(f"  File written: {wdq.fileWritten}")

        # Get time data
        time_data = wdq.time()
        logger.info(f"Time range: {time_data[0]:.6f} to {time_data[-1]:.6f} seconds")

        # Create DataFrame with time column
        df_data = {'Time_seconds': time_data}

        # Add data for each channel
        for ch in range(1, wdq.nChannels + 1):
            channel_data = wdq.data(ch)
            unit = wdq.unit(ch)
            annotation = wdq.chAnnotation(ch)

            # Create descriptive column name
            if annotation and annotation.strip():
                col_name = f"Channel_{ch}_{annotation.strip()}"
            else:
                col_name = f"Channel_{ch}"

            if unit and unit.strip():
                col_name += f"_{unit.strip()}"

            df_data[col_name] = channel_data

            logger.info(f"Channel {ch}: {annotation} ({unit}) - {len(channel_data)} samples")

        # Create DataFrame
        df = pd.DataFrame(df_data)

        # If there's only one data channel, rename it to "Current" for convenience
        if wdq.nChannels == 1:
            data_columns = [col for col in df.columns if col != 'Time_seconds']
            if data_columns:
                df = df.rename(columns={data_columns[0]: 'Current'})
                logger.info("Renamed single data channel to 'Current'")

        logger.info(f"Created DataFrame with {len(df)} rows and {len(df.columns)} columns")
        logger.info(f"Columns: {list(df.columns)}")

        # Save to Excel
        logger.info(f"Saving to Excel file: {excel_file_path}")
        df.to_excel(excel_file_path, index=False, engine='openpyxl')

        logger.info(f"Successfully converted '{wdq_file_path}' to '{excel_file_path}'")
        return str(excel_file_path)

    except Exception as e:
        logger.error(f"Error converting WinDaq file: {str(e)}")
        raise


def main():
    """Main function to handle command line arguments and file conversion."""
    parser = argparse.ArgumentParser(
        description="Convert WinDaq files to Excel format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s data.csv                    # Convert CSV to Excel
  %(prog)s data.wdq output.xlsx        # Convert WinDaq to Excel
  %(prog)s "my file.csv"               # Handle files with spaces
        """
    )
    
    parser.add_argument('input_file', help='Input file path (.csv, .wdq, .wdh, or .wdc)')
    parser.add_argument('output_file', nargs='?', help='Output Excel file path (optional)')
    parser.add_argument('--force-csv', action='store_true', 
                       help='Force treating input as CSV file')
    parser.add_argument('--force-windaq', action='store_true',
                       help='Force treating input as WinDaq file')
    
    args = parser.parse_args()
    
    # Check if input file exists
    if not os.path.exists(args.input_file):
        logger.error(f"Input file not found: {args.input_file}")
        sys.exit(1)
    
    # Determine file type and conversion method
    input_path = Path(args.input_file)
    file_extension = input_path.suffix.lower()
    
    try:
        if args.force_csv or file_extension == '.csv':
            # Convert CSV to Excel
            output_file = convert_csv_to_excel(args.input_file, args.output_file)
        elif args.force_windaq or file_extension in ['.wdq', '.wdh', '.wdc']:
            # Convert WinDaq file directly
            output_file = convert_windaq_direct(args.input_file, args.output_file)
        else:
            logger.error(f"Unsupported file type: {file_extension}")
            logger.info("Supported types: .csv, .wdq, .wdh, .wdc")
            sys.exit(1)
        
        logger.info(f"Conversion completed successfully!")
        logger.info(f"Output file: {output_file}")
        
    except Exception as e:
        logger.error(f"Conversion failed: {str(e)}")
        logger.info("\nRecommended workflow:")
        logger.info("1. Open your WinDaq file in WinDaq Waveform Browser")
        logger.info("2. Go to File > Save As...")
        logger.info("3. Choose 'Spreadsheet print (CSV)' format")
        logger.info("4. Save as .csv file")
        logger.info("5. Run this script on the CSV file")
        sys.exit(1)


if __name__ == "__main__":
    main()
