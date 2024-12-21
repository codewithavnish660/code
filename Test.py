import pandas as pd
from opencage.geocoder import OpenCageGeocode

# Initialize OpenCage API
API_KEY = '5c845f7c709c4a5ba85ccd437c0c171b'  # Replace with your OpenCage API key
geocoder = OpenCageGeocode(API_KEY)

# Load the Excel file
input_file = 'sample_data.xlsx'  # Replace with your file name
output_file = 'Geocoordinates.xlsx'

# Load the data
df = pd.read_excel(input_file, sheet_name='Sheet1')

# Define a function to get latitude and longitude
def get_lat_lon(address):
    try:
        result = geocoder.geocode(address)
        if result:
            return result[0]['geometry']['lat'], result[0]['geometry']['lng']
    except Exception as e:
        print(f"Error geocoding {address}: {e}")
    return None, None

# Combine address columns to form a full address (convert to string)
df['full_address'] = df[['c_add1', 'c_add2', 'c_add3']].fillna('').astype(str).agg(', '.join, axis=1)

# Fetch latitude and longitude
latitudes = []
longitudes = []

for address in df['full_address']:
    lat, lon = get_lat_lon(address)
    latitudes.append(lat)
    longitudes.append(lon)

# Add latitude and longitude to the dataframe
df['latitude'] = latitudes
df['longitude'] = longitudes

# Save the updated data to a new Excel file
df.to_excel(output_file, index=False)

print(f"File saved as {output_file}")
