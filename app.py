from flask import Flask, render_template, request, redirect, send_from_directory, abort
import os
from werkzeug.utils import secure_filename
import config

app = Flask(__name__)

"""
string:
int:
float:
path:
uuid:
"""

def allowed_excel(filename):
    if not "." in filename:
        return False
    ext = filename.rsplit(".", 1)[1]
    if ext.upper() in config.ALLOWED_EXTENSIONS:
        return True
    else:
        return False

@app.route('/',  methods=["GET", "POST"])
def upload_excel():
    if request.method == "POST":
        if request.files["target_excel"]:
            target_excel = request.files["target_excel"]
            if target_excel.filename == "":
                print("Must have a filename")
                return redirect(request.url)
            if not allowed_excel(target_excel.filename):
                print("File extension is not allowed")
                return redirect(request.url)
            else:
                filename = secure_filename(target_excel.filename)
                target_excel.save(os.path.join(config.EXCEL_UPLOADS, filename))
            print("Target Excel File Saved")

            if request.files["sals_excel"]:
                sals_excel = request.files["sals_excel"]
                if sals_excel.filename == "":
                    print("Must have a filename")
                    return redirect(request.url)
                if not allowed_excel(sals_excel.filename):
                    print("File extension is not allowed")
                    return redirect(request.url)
                else:
                    filename = secure_filename(sals_excel.filename)
                    sals_excel.save(os.path.join(config.EXCEL_UPLOADS, filename))
                print("Salsify Excel File Saved")
            return redirect(request.url)
    return render_template("public/templates/upload_excel.html")

@app.route('/get-excel/<excel_download>')
def get_excel(excel_download):
    try:
        return send_from_directory(directory=config.CLIENT_EXCELS, filename=excel_download, as_attachment=True, path='/')
    except FileNotFoundError:
        abort(404)
    return "Ready for download"

if __name__ == '__main__':
    app.run()
