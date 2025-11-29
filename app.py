from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import os

from extractor import run_extraction

app = Flask(__name__)
app.secret_key = "some-secret-key-change-this"  # needed for flash messages


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        folder = request.form.get("folder_path", "").strip()

        if not folder:
            flash("Please enter a folder path.", "error")
            return redirect(url_for("index"))

        try:
            output_file = run_extraction(folder)
        except Exception as e:
            flash(f"Error: {e}", "error")
            return redirect(url_for("index"))

        # send the Excel file to browser as download
        return send_file(
            output_file,
            as_attachment=True,
            download_name=os.path.basename(output_file)
        )

    return render_template("index.html")
    

if __name__ == "__main__":
    # debug=True is optional; you can set to False in production
    app.run(host="127.0.0.1", port=5000, debug=True)
