from flask import Flask, request, jsonify, send_file
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from io import BytesIO

app = Flask(__name__)

SECRET_KEY = "Rinben@123$$"

# ----------------- FAST HAVERSINE FUNCTION (Vectorized) -----------------
def haversine(lat1, lon1, lat2, lon2):
    R = 6371000  # Earth radius in meters
    lat1, lon1, lat2, lon2 = map(np.radians, [lat1, lon1, lat2, lon2])

    dlat = lat2 - lat1
    dlon = lon2 - lon1

    a = np.sin(dlat/2)**2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon/2)**2
    c = 2 * np.arcsin(np.sqrt(a))
    return R * c

# ----------------- FAST NEAREST NEIGHBOR ORDERING -----------------
def fast_reorder(df):
    n = len(df)
    visited = np.zeros(n, dtype=bool)
    order = [0]  # Start from first point
    visited[0] = True

    lat = df["Latitude"].values
    lon = df["Longitude"].values

    for _ in range(1, n):
        last = order[-1]

        dist_all = haversine(lat[last], lon[last], lat, lon)

        # Ignore visited rows
        dist_all[visited] = np.inf  

        next_point = np.argmin(dist_all)
        order.append(next_point)
        visited[next_point] = True

    df = df.iloc[order].reset_index(drop=True)
    df["Sl_No"] = np.arange(1, n + 1)
    return df


@app.route("/process", methods=["POST"])
def process_file():

    api_key = request.headers.get("x-api-key")
    if api_key != SECRET_KEY:
        return jsonify({"error": "Unauthorized"}), 401

    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    filename = file.filename

    df = pd.read_csv(file)
    df.columns = df.columns.str.strip()

    # Extract first 5 columns
    df = df.iloc[:, :5]
    df.columns = ["ObjectId", "Latitude", "Longitude", "Altitude", "Name"]

    # Sort initially
    df = df.sort_values("ObjectId").reset_index(drop=True)

    # Optimized ordering
    df = fast_reorder(df)

    # Compute distances (vectorized)
    lat = df["Latitude"].values
    lon = df["Longitude"].values

    distances = np.zeros(len(df))
    distances[1:] = haversine(lat[:-1], lon[:-1], lat[1:], lon[1:])
    df["Distance [m]"] = np.round(distances, 2)

    df["Remarks"] = ""

    # ---------------- FORCE EXACT COLUMN ORDER ----------------
    df = df.reindex(columns=[
        "ObjectId",
        "Sl_No",
        "Latitude",
        "Longitude",
        "Altitude",
        "Name",
        "Distance [m]",
        "Remarks"
    ])

    # ---------- Export Excel ----------
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    wb = load_workbook(output)
    ws = wb.active

    tab = Table(displayName="DGPS_Table", ref=f"A1:H{len(df) + 1}")
    style = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    tab.tableStyleInfo = style
    ws.add_table(tab)
    ws.row_dimensions[1].height = 30

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    return send_file(
        final_output,
        download_name=filename.replace(".csv", "_CLEAN.xlsx"),
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    app.run(debug=True)
