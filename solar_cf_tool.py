# solar_cf_tool.py

import streamlit as st
import pandas as pd
import pvlib
import matplotlib.pyplot as plt
import folium
from streamlit_folium import st_folium
import time
from datetime import timedelta
import branca
import numpy as np
import io

# --- App Settings ---
st.set_page_config(page_title="Solar Specific Production & CF Estimator", layout="wide")

# --- Title and Instructions ---
st.title("☀️ Solar Specific Production & Capacity Factor (CF) Tool")
st.markdown("Upload a site list and estimate Specific Production (kWh/kWp) and Capacity Factor (%).")

# --- Upload File ---
uploaded_file = st.file_uploader("📁 Upload your 'sites.xlsx' file (columns: name, Lat, Long)", type=["xlsx"])

# --- Sidebar Parameters ---
st.sidebar.header("Parameters")
tilt_angle = st.sidebar.slider("Tilt Angle (°)", min_value=0, max_value=60, value=20)
PR = st.sidebar.slider("Performance Ratio (PR)", min_value=0.5, max_value=1.0, value=0.85)

# --- Initialize session state ---
if 'analysis_done' not in st.session_state:
    st.session_state['analysis_done'] = False
if 'result_df' not in st.session_state:
    st.session_state['result_df'] = None

# --- If file is uploaded ---
if uploaded_file:
    site_df = pd.read_excel(uploaded_file)
    st.success(f"✅ File uploaded successfully. {len(site_df)} sites found.")

    run_button = st.button("🚀 Run Analysis")

    if run_button:
        st.session_state['analysis_done'] = False  # Reset first
        start_time = time.time()
        results = []
        progress_bar = st.progress(0)

        # --- Loop Through Sites ---
        for idx, row in site_df.iterrows():
            name = row['name']
            lat = row['Lat']
            lon = row['Long']
            try:
                tmy_data = pvlib.iotools.get_pvgis_tmy(
                    latitude=lat,
                    longitude=lon,
                    outputformat='json',
                    usehorizon=True,
                    startyear=2005,
                    endyear=2023,
                    map_variables=True,
                    url='https://re.jrc.ec.europa.eu/api/',
                    timeout=30
                )[0]

                solar_pos = pvlib.solarposition.get_solarposition(tmy_data.index, lat, lon)
                poa_irradiance = pvlib.irradiance.get_total_irradiance(
                    surface_tilt=tilt_angle,
                    surface_azimuth=180,
                    dni=tmy_data['dni'],
                    ghi=tmy_data['ghi'],
                    dhi=tmy_data['dhi'],
                    solar_zenith=solar_pos['zenith'],
                    solar_azimuth=solar_pos['azimuth']
                )

                annual_gii = poa_irradiance['poa_global'].sum() / 1000
                annual_ghi = tmy_data['ghi'].sum() / 1000
                specific_prod = annual_gii * PR
                uplift = ((annual_gii / annual_ghi) - 1) * 100
                cf_percent = round((specific_prod / 8760) * 100, 3)

                results.append({
                    'Site': name,
                    'Lat': lat,
                    'Long': lon,
                    'GHI (kWh/m²)': round(annual_ghi, 1),
                    'GII (kWh/m²)': round(annual_gii, 1),
                    'Uplift (%)': round(uplift, 1),
                    'Specific Production (kWh/kWp)': round(specific_prod, 1),
                    'CF (%)': cf_percent
                })

            except Exception as e:
                results.append({
                    'Site': name,
                    'Lat': lat,
                    'Long': lon,
                    'GII (kWh/m²)': None,
                    'Specific Production (kWh/kWp)': None,
                    'CF (%)': None
                })
                st.error(f"❌ Failed for {name}: {e}")

            # Update progress bar
            progress = (idx + 1) / len(site_df)
            progress_bar.progress(progress)

        # --- Save results to session ---
        result_df = pd.DataFrame(results)
        st.session_state['result_df'] = result_df
        
        st.session_state['analysis_done'] = True

        elapsed = timedelta(seconds=int(time.time() - start_time))
        st.success(f"✅ Analysis completed! (⏱️ {elapsed})")
        st.dataframe(result_df)

# --- Show results if analysis was done ---
if st.session_state['analysis_done'] and st.session_state['result_df'] is not None:
    result_df = st.session_state['result_df']
    st.dataframe(result_df)

    # --- Download Button ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False)
    data = output.getvalue()

    st.download_button(
        label="📥 Download Results as Excel",
        data=data,
        file_name="solar_cf_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # --- Plots ---
    valid_df = result_df.dropna(subset=['CF (%)'])
    cf_values = valid_df['CF (%)']

    st.subheader("📈 Capacity Factor Analysis")

    col1, col2 = st.columns(2)
    with col1:
        fig1, ax1 = plt.subplots()
        ax1.hist(cf_values, bins=10, color='skyblue', edgecolor='black')
        ax1.set_title("CF Histogram")
        ax1.set_xlabel("CF (%)")
        ax1.set_ylabel("Number of Sites")
        st.pyplot(fig1)

    with col2:
        fig2, ax2 = plt.subplots()
        ax2.boxplot(cf_values, vert=False, patch_artist=True,
                    boxprops=dict(facecolor='lightgreen'))
        ax2.set_title("CF Boxplot")
        ax2.set_xlabel("CF (%)")
        st.pyplot(fig2)

    fig3, ax3 = plt.subplots(figsize=(10, len(valid_df) * 0.4))
    sorted_cf_df = valid_df.sort_values('CF (%)', ascending=False)
    ax3.barh(sorted_cf_df['Site'], sorted_cf_df['CF (%)'], color='salmon', edgecolor='black')
    ax3.set_xlabel("CF (%)")
    ax3.set_title("Capacity Factor by Site")
    ax3.invert_yaxis()
    ax3.grid(axis='x')
    st.pyplot(fig3)

    # --- Interactive Map ---
    st.subheader("🗺️ Site Map with Satellite Basemap")

    map_center = [valid_df['Lat'].mean(), valid_df['Long'].mean()]
    site_map = folium.Map(location=map_center, zoom_start=5, tiles=None)

    folium.TileLayer(
        tiles='https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}',
        attr='Google',
        name='Google Satellite',
        overlay=False,
        control=True
    ).add_to(site_map)

    min_cf = valid_df['CF (%)'].min()
    max_cf = valid_df['CF (%)'].max()
    color_scale = branca.colormap.LinearColormap(
        colors=['red', 'orange', 'yellow', 'green'],
        vmin=min_cf, vmax=max_cf,
        caption='Capacity Factor (%)'
    )

    for _, row in valid_df.iterrows():
        folium.CircleMarker(
            location=[row['Lat'], row['Long']],
            radius=6,
            color=color_scale(row['CF (%)']),
            fill=True,
            fill_color=color_scale(row['CF (%)']),
            fill_opacity=0.8,
            popup=folium.Popup(f"{row['Site']}<br>CF: {row['CF (%)']:.2f}%", max_width=200)
        ).add_to(site_map)

    color_scale.add_to(site_map)
    st_data = st_folium(site_map, width=900)
