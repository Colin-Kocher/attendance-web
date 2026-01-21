from flask import Flask, render_template, request, send_file
import pandas as pd
from datetime import time
import io
import os

app = Flask(__name__)

# =============================
# CONFIG (EXACT MATCH)
# =============================
TIMESTAMP_COL = "event.published"
NAME_COL = "actor.display_name"

LATE_ENTRY_TIME = time(10, 0)      # 10:00 AM
EXPECTED_EXIT_TIME = time(18, 0)   # 6:00 PM


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")

        if not file:
            return render_template("index.html", error="No file selected")

        try:
            # Load file
            if file.filename.lower().endswith(".csv"):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file)

            # Process data (IDENTICAL logic)
            df[TIMESTAMP_COL] = pd.to_datetime(df[TIMESTAMP_COL])

            df["Date"] = df[TIMESTAMP_COL].dt.date
            df["Time"] = df[TIMESTAMP_COL].dt.time

            summary = (
                df.groupby([NAME_COL, "Date"])
                .agg(
                    In_Time=("Time", "min"),
                    Out_Time=("Time", "max"),
                    In_Datetime=(TIMESTAMP_COL, "min"),
                    Out_Datetime=(TIMESTAMP_COL, "max"),
                )
                .reset_index()
            )

            summary["Total_Hours"] = (
                summary["Out_Datetime"] - summary["In_Datetime"]
            ).dt.total_seconds() / 3600
            summary["Total_Hours"] = summary["Total_Hours"].round(2)

            summary["Late_Entry"] = summary["In_Time"].apply(
                lambda t: "Yes" if t > LATE_ENTRY_TIME else "No"
            )

            summary["Early_Exit"] = summary["Out_Time"].apply(
                lambda t: "Yes" if t < EXPECTED_EXIT_TIME else "No"
            )

            final_df = summary[
                [
                    NAME_COL,
                    "Date",
                    "In_Time",
                    "Out_Time",
                    "Total_Hours",
                    "Late_Entry",
                    "Early_Exit",
                ]
            ].rename(
                columns={
                    NAME_COL: "Name",
                    "In_Time": "In",
                    "Out_Time": "Out",
                }
            )

            final_df = final_df.sort_values(["Name", "Date"])

            # Save Excel in memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final_df.to_excel(writer, index=False)

            output.seek(0)

            base_name = os.path.splitext(file.filename)[0]
            output_name = f"{base_name}_attendance_summary.xlsx"

            return send_file(
                output,
                as_attachment=True,
                download_name=output_name,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            return render_template("index.html", error=str(e))

    return render_template("index.html")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)

