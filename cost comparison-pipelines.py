# -*- coding: utf-8 -*-
"""
Created on Tue Jul 22 11:10:45 2025

@author: sarla
"""

import pandas as pd
import plotly.express as px
import plotly.io as pio
import webbrowser


# Load Excel data
file_path = "C:/Users/sarla/Downloads/Hydrogen economics-Solomon (1).xlsx"

# Pipeline data starts after row 57
pipeline_df = pd.read_excel(
    file_path,
    sheet_name="Pipeline Model",
    skiprows=57
)

# Truck data starts after row 17 — actual headers are in row 18
truck_df = pd.read_excel(
    file_path,
    sheet_name="Trucks Model",
    skiprows=17,
    header=0
)

# Strip whitespace from column headers
pipeline_df.columns = pipeline_df.columns.str.strip()
truck_df.columns = truck_df.columns.str.strip()

# Rename columns for consistency
pipeline_df = pipeline_df.rename(columns={
    "Diameter of pipeline(mm)": "Diameter_mm",
    "Daily Hydrogen Demand{mH2}(kg/day)": "Hydrogen_Demand",
    "Length of pipeline {L}(km)": "Distance_km",
    "LCOHT,pipeline": "LCOHT"
})

truck_df = truck_df.rename(columns={
    "Component": "Component",
    "Daily Hydrogen Demand{mH2}(kg/day)": "Hydrogen_Demand",
    "Distance{d}(km)": "Distance_km",
    "Total LCOHT for compressed hydrogen gas via trucks and trailers (€/kg)": "LCOHT"
})

# Filter for 30,000 kg/day demand
target_demand = 30000

pipeline_filtered = pipeline_df[pipeline_df["Hydrogen_Demand"] == target_demand]
truck_filtered = truck_df[truck_df["Hydrogen_Demand"] == target_demand]

# Select specific pipeline diameters
pipeline_final = pipeline_filtered[pipeline_filtered["Diameter_mm"].isin([100, 150, 200])].copy()
pipeline_final["Component"] = pipeline_final["Diameter_mm"].astype(str) + " mm (Pipeline)"

# Select specific trailer types
truck_final = truck_filtered[truck_filtered["Component"].isin(["350 bar (Trailer)", "540 bar (Trailer)"])].copy()

# Combine into one DataFrame
combined_df = pd.concat([
    truck_final[["Distance_km", "LCOHT", "Component"]],
    pipeline_final[["Distance_km", "LCOHT", "Component"]]
])

# Only show data at 25, 100, 250, 500 km
filtered_df = combined_df[combined_df["Distance_km"].isin([25, 100, 250, 500])]

# Plot bar chart
fig = px.bar(
    filtered_df,
    x="Distance_km",
    y="LCOHT",
    color="Component",
    barmode="group",
    title=f"LCOHT vs Distance for Hydrogen Transport (at {target_demand} kg/day)",
    labels={"Distance_km": "Distance (km)", "LCOHT": "LCOHT (€/kg)", "Component": "Transport Mode"}
)

# Show only specified x-ticks
fig.update_xaxes(tickvals=[25, 100, 250, 500])

# Show the chart in a new browser tab
fig.show()
fig.write_html("lcoht_all_components.html")
webbrowser.open("lcoht_all_components.html")
fig.write_image("lcoht_all_components.png")
print("✅ Chart generated with: 100mm, 150mm, 200mm pipelines + 350 bar & 540 bar trailers.")
