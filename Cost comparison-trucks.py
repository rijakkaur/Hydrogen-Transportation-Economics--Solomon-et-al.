import pandas as pd
import plotly.express as px
import webbrowser

# Load Excel data (Trucks model)
file_path = "C:/Users/sarla/Downloads/Hydrogen economics-Solomon (1).xlsx"
truck_df = pd.read_excel(file_path, sheet_name="Trucks Model", skiprows=17)

# Strip whitespace from column names
truck_df.columns = truck_df.columns.str.strip()

# Rename necessary columns for clarity
truck_df = truck_df.rename(columns={
    "Component": "Component",
    "Daily Hydrogen Demand{mH2}(kg/day)": "Hydrogen_Demand",
    "Distance{d}(km)": "Distance_km",
    "Total LCOHT for compressed hydrogen gas via trucks and trailers (€/kg)": "LCOHT"
})

# Filter for 30,000 kg/day demand
truck_filtered = truck_df[truck_df["Hydrogen_Demand"] == 30000]

# Filter only for 350 and 540 bar trailers
selected_components = ["Trailer(350 bar)", "Trailer(540 bar)"]
truck_filtered = truck_filtered[truck_filtered["Component"].isin(selected_components)]

# Keep only exact distances
target_distances = [25, 100, 250, 500]
truck_filtered = truck_filtered[truck_filtered["Distance_km"].isin(target_distances)]

# Update labels for display
truck_filtered["Component"] = truck_filtered["Component"].str.replace("Trailer", "Trailer ", regex=False)

# Plot bar chart
fig = px.bar(
    truck_filtered,
    x="Distance_km",
    y="LCOHT",
    color="Component",
    barmode="group",
    title="LCOHT vs Distance for Hydrogen Transport by Truck (30,000 kg/day)",
    labels={
        "Distance_km": "Distance (km)",
        "LCOHT": "LCOHT (€/kg)",
        "Component": "Trailer Type"
    }
)

# Only show tick values for the four distances
fig.update_xaxes(tickvals=target_distances)

# Show the chart in browser and save
fig.write_html("lcoht_truck_comparison.html")
webbrowser.open("lcoht_truck_comparison.html")
fig.write_image("lcoht_truck_comparison.png")
print("✅ Chart generated for 350 vs 540 bar trailers at 30,000 kg/day.")
