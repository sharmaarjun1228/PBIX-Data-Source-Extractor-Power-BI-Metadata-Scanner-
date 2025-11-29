from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import os
import tempfile
from extractor import run_extraction

app = Flask(__name__)
app.secret_key = "my-secret"

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        files = request.files.getlist("pbix_files")

        if not files:
            flash("Please select a folder containing PBIX files.", "error")
            return redirect(url_for("index"))

        # Create temporary folder
        temp_dir = tempfile.mkdtemp()

        # Save uploaded files
        for f in files:
            filename = f.filename
            # Only save pbix files
            if filename.lower().endswith(".pbix"):
                save_path = os.path.join(temp_dir, os.path.basename(filename))
                f.save(save_path)

        try:
            output_file = run_extraction(temp_dir)
        except Exception as e:
            flash(f"Error: {e}", "error")
            return redirect(url_for("index"))

        return send_file(
            output_file,
            as_attachment=True,
            download_name=os.path.basename(output_file)
        )

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
