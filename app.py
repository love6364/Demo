from flask import Flask, render_template, request, send_file
from io import BytesIO
from datetime import date

from utils.converter import *

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def home():
    if request.method == "POST":
        try:
            master = request.files['master']
            party = request.files['party']

            mdf = load_master(master.read())
            pdf = load_party(party.read())

            col_map = map_columns(pdf)

            result = convert(mdf, pdf, col_map, {})

            excel = to_excel(result)

            return send_file(BytesIO(excel),
                             download_name="output.xlsx",
                             as_attachment=True)

        except Exception as e:
            import traceback
            return f"<pre>{traceback.format_exc()}</pre>"

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)