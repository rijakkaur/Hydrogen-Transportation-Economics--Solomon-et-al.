import pandas as pd
import plotly.graph_objects as go
import webbrowser

# Load Excel file and read specific data sections
pipeline_df = pd.read_excel(
    "C:/Users/sarla/Downloads/Hydrogen economics-Solomon (1).xlsx",
    sheet_name="Pipeline Model",
    skiprows=57  # Data starts from row 58 (index 57)
)

truck_df = pd.read_excel(
    "C:/Users/sarla/Downloads/Hydrogen economics-Solomon (1).xlsx",
    sheet_name="Trucks Model",
    skiprows=17  # Data starts from row 18 (index 17)
)

# Clean column names
pipeline_df.columns = pipeline_df.columns.str.strip()
truck_df.columns = truck_df.columns.str.strip()

# Rename key columns
pipeline_df = pipeline_df.rename(columns={
    "Diameter of pipeline(mm)": "Diameter_mm",
    "Daily Hydrogen Demand{mH2}(kg/day)": "Hydrogen_Demand",
    "Length of pipeline {L}(km)": "Length_km",
    "LCOHT,pipeline": "LCOHT"
})

truck_df = truck_df.rename(columns={
    "Component": "Component",
    "Daily Hydrogen Demand{mH2}(kg/day)": "Hydrogen_Demand",
    "Distance{d}(km)": "Length_km",
    "Total LCOHT for compressed hydrogen gas via trucks and trailers (€/kg)": "LCOHT"
})

# Initialize 3D figure
fig = go.Figure()

# --- Add pipeline surfaces ---
pipeline_df["Diameter_mm"] = pd.to_numeric(pipeline_df["Diameter_mm"], errors='coerce')
for diameter in sorted(pipeline_df["Diameter_mm"].dropna().unique()):

    temp = pipeline_df[pipeline_df["Diameter_mm"] == diameter]
    pivot = temp.pivot_table(index="Hydrogen_Demand", columns="Length_km", values="LCOHT")
    
    fig.add_trace(go.Surface(
        z=pivot.values,
        x=pivot.index,
        y=pivot.columns,
        name=f"{int(diameter)} mm (Pipeline)",
        showscale=False,
        colorscale='Viridis',
        opacity=0.85,
        showlegend=True
    ))

# --- Add truck surfaces ---
for component in truck_df["Component"].dropna().unique():
    temp = truck_df[truck_df["Component"] == component]
    pivot = temp.pivot_table(index="Hydrogen_Demand", columns="Length_km", values="LCOHT")
    
    fig.add_trace(go.Surface(
        z=pivot.values,
        x=pivot.index,
        y=pivot.columns,
        name=f"{component} (Trailer)",
        showscale=False,
        colorscale='Reds',
        opacity=0.85,
        showlegend=True
    ))

# Layout customization
fig.update_layout(
    title="LCOHT vs Hydrogen Demand and Transport Distance",
    scene=dict(
        xaxis_title="Hydrogen Demand (kg/day)",
        yaxis_title="Transport Distance (km)",
        zaxis_title="LCOHT (€/kg)",
        xaxis=dict(tickformat=".0f"),
        yaxis=dict(tickformat=".0f"),
        zaxis=dict(tickformat=".3f")
    ),
    width=1100,
    height=850,
    template="plotly_white",
    legend=dict(
        title="Transport Option",
        x=0.75,
        y=0.95,
        bgcolor="rgba(255,255,255,0.5)"
    )
)

# Save and open in browser
fig.write_html("hydrogen_transport_cost_comparison.html")
webbrowser.open("hydrogen_transport_cost_comparison.html")
fig.show()
