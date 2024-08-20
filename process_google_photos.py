import os
import sys
import zipfile
import json
import time
import shutil
from datetime import datetime
import piexif
import pywintypes
import win32file
import win32con

def set_file_creation_time(file_path, timestamp):
    # Convert timestamp to a format Windows can use
    creation_time = pywintypes.Time(datetime.fromtimestamp(int(timestamp)))
    
    # Open the file and set its creation time using the Windows API
    handle = win32file.CreateFile(
        file_path, win32con.GENERIC_WRITE, 0, None,
        win32con.OPEN_EXISTING, win32con.FILE_ATTRIBUTE_NORMAL, None
    )
    win32file.SetFileTime(handle, creation_time, None, None)
    handle.close()

def unzip_and_set_times(zip_path):
    extract_path = os.path.join(os.path.dirname(zip_path), 'Google_Photos_Extracted')
    
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        for file_info in zip_ref.infolist():
            # Extract file
            extracted_path = zip_ref.extract(file_info, extract_path)

            # Skip directories
            if os.path.isdir(extracted_path):
                continue

            # Look for the corresponding JSON file
            json_path = os.path.join(extract_path, file_info.filename + '.json')
            if os.path.exists(json_path):
                with open(json_path, 'r') as f:
                    metadata = json.load(f)
                
                if 'photoTakenTime' in metadata and 'timestamp' in metadata['photoTakenTime']:
                    # Set file creation time
                    set_file_creation_time(extracted_path, metadata['photoTakenTime']['timestamp'])

    return extract_path

def update_exif_data(file_path, geo_data):
    # Load the EXIF data from the file
    exif_dict = piexif.load(file_path)
    
    # Convert latitude and longitude to DMS format (degrees, minutes, seconds)
    def to_dms(value):
        degrees = int(value)
        minutes = int((value - degrees) * 60)
        seconds = int((value - degrees - minutes/60) * 3600 * 10000)
        return ((degrees, 1), (minutes, 1), (seconds, 10000))
    
    lat_dms = to_dms(geo_data['latitude'])
    lon_dms = to_dms(geo_data['longitude'])
    
    # Update GPS data in EXIF
    exif_dict['GPS'][piexif.GPSIFD.GPSLatitude] = lat_dms
    exif_dict['GPS'][piexif.GPSIFD.GPSLongitude] = lon_dms
    exif_dict['GPS'][piexif.GPSIFD.GPSLatitudeRef] = 'N' if geo_data['latitude'] >= 0 else 'S'
    exif_dict['GPS'][piexif.GPSIFD.GPSLongitudeRef] = 'E' if geo_data['longitude'] >= 0 else 'W'
    
    if 'altitude' in geo_data:
        exif_dict['GPS'][piexif.GPSIFD.GPSAltitude] = (int(geo_data['altitude'] * 1000), 1000)
        exif_dict['GPS'][piexif.GPSIFD.GPSAltitudeRef] = 0 if geo_data['altitude'] >= 0 else 1
    
    # Save updated EXIF data back to the file
    exif_bytes = piexif.dump(exif_dict)
    piexif.insert(exif_bytes, file_path)

def process_json_files(extract_path):
    # Walk through the directories and process each JSON file
    for root, dirs, files in os.walk(extract_path):
        for file in files:
            if file.endswith('.json'):
                json_path = os.path.join(root, file)
                media_file = json_path[:-5]  # Remove the '.json' extension to get the media file
                
                if os.path.exists(media_file):
                    # Read the JSON metadata
                    with open(json_path, 'r') as f:
                        metadata = json.load(f)
                    
                    # Update the EXIF GPS data
                    if 'geoData' in metadata:
                        update_exif_data(media_file, metadata['geoData'])

def main():
    if len(sys.argv) != 2:
        print("Usage: python script.py <path_to_zip_file>")
        sys.exit(1)
    
    zip_path = sys.argv[1]
    
    if not os.path.exists(zip_path):
        print(f"File not found: {zip_path}")
        sys.exit(1)
    
    # Unzip and set creation times
    extract_path = unzip_and_set_times(zip_path)
    
    # Process JSON files for EXIF updates
    process_json_files(extract_path)

    print("Processing complete.")

if __name__ == "__main__":
    main()
