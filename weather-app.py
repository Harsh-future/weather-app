import streamlit as st
import pandas as pd
import requests
import json
import io

ASHRAE_API_URL = "https://ashrae-meteo.info/v2.0/request_places.php"
ASHRAE_WEATHER_API_URL = "https://ashrae-meteo.info/v2.0/request_meteo_parametres.php"

def fetch_station(lat, lon):
    """Fetch nearest ASHRAE station based on latitude & longitude."""
    request_params = {
        "lat": lat,
        "long": lon,
        "number": "10",  # Fetch only 1 nearest station
        "ashrae_version": "2017"
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Accept": "application/json,text/html,application/xhtml+xml",
        "Referer": "http://google.com",
    }
    
    try:
        resp = requests.post(ASHRAE_API_URL, data=request_params,headers=headers)
        resp_json = resp.json()
        stations = resp_json.get("meteo_stations", [])
        return stations[0] if stations else None

    except requests.exceptions.RequestException as e:
        st.error(f"❌ Error fetching station: {e}")
        return None

def fetch_weather_data(station_data):
    """Fetch weather data for the given ASHRAE station."""
    if not station_data:
        print("❌ No station data received.")
        return None

    request_params = {
        "wmo": station_data.get("wmo"),
        "ashrae_version": "2017",
        "si_ip": "SI"
    }
    url = "https://ashrae-meteo.info/v2.0/request_meteo_parametres.php"

    try:
        print(f"🔍 Fetching weather data for WMO: {station_data.get('wmo')}")
        print(f"🔍 Requesting: {url} with {request_params}")
        
        resp = requests.post(url, data=request_params)
        print("🔍 API Status Code:", resp.status_code)
        print("🔍 Response Content-Type:", resp.headers.get("Content-Type"))
        print("🔍 API Response (First 500 chars):", resp.text[:500])

        if resp.status_code != 200:
            print(f"❌ Error: Received status code {resp.status_code}")
            return None

        if not resp.text.strip():  
            print("❌ Error: Empty response from API.")
            return None

        resp_json = json.loads(resp.content.decode("utf-8-sig"))
        stations = resp_json.get('meteo_stations', [])

        if not stations:
            print("❌ No weather stations found in API response.")
            return None

        station = stations[0]
        weather_data = {
            "cooling_DB_MCWB_0.4_DB": station.get("cooling_DB_MCWB_0.4_DB", "n/a"),
            "cooling_DB_MCWB_0.4_MCWB": station.get("cooling_DB_MCWB_0.4_MCWB", "n/a")
        }
        print(f"✅ Weather Data Fetched: {weather_data}")
        return weather_data

    except requests.exceptions.RequestException as e:
        print(f"❌ API Request Error: {e}")
        return None
    except json.JSONDecodeError as e:
        print(f"❌ JSON Decode Error: {e}")
        print("🔍 Raw API Response:", resp.text)
        return None
    
st.title("📊 Outdoor Condition Data")
st.write("Upload an Excel file with Latitude and Longitude columns, and get updated values for DB and MCWB.")

uploaded_file = st.file_uploader("📤 Upload Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    
    if "Latitude" not in df.columns or "Longitude" not in df.columns:
        st.error("❌ Excel file must contain 'Latitude' and 'Longitude' columns.")
    else:
        if st.button("🔄 Process Data"):
            with st.spinner("Fetching data..."):
                df["DB"] = ""
                df["MCWB"] = ""

                for i, row in df.iterrows():
                    lat, lon = row["Latitude"], row["Longitude"]
                    station = fetch_station(lat, lon)
                    weather_data = fetch_weather_data(station)

                    if weather_data:
                        df.at[i, "DB"] = weather_data["cooling_DB_MCWB_0.4_DB"]
                        df.at[i, "MCWB"] = weather_data["cooling_DB_MCWB_0.4_MCWB"]
                    else:
                        df.at[i, "DB"] = "n/a"
                        df.at[i, "MCWB"] = "n/a"

                st.success("✅ Data processing complete!")

            # Convert DataFrame to Excel file
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Updated Data")
            output.seek(0)

            st.download_button(
                label="📥 Download Updated Excel",
                data=output,
                file_name="updated_weather_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
