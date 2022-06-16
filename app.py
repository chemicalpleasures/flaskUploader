from flask import Flask, render_template, request, redirect, send_from_directory, abort
import os
from werkzeug.utils import secure_filename
import config
import pandas as pd
import openpyxl

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

@app.route('/parser', methods=['POST', 'GET'])
def parser():
    if request.method == "POST":
        # defines datasets. wb is the list of Target errors from the target portal. wb2 is product data from salsify.
        wb = 'static/excel_files/uploads/target.xlsx'
        wb2 = 'static/excel_files/uploads/inv.xlsx'

        df1 = pd.read_excel(wb)
        df2 = pd.read_excel(wb2)

        # renames first column
        df1.rename(columns={"Partner SKU": "Inventory Number"}, inplace=True)

        merge = pd.merge(df1, df2, how="left", on="Inventory Number")

        # removes unwanted columns
        merge.drop(['Barcode', 'TCIN', 'Published', 'Error Code'], axis=1, inplace=True)

        # creates condition which removes all product statuses and reasons we don't want
        cond1 = (merge["Product Status (Computed)"] == "Resourcing - Quality/Defective") | (
                    merge["Product Status (Computed)"] == "Resourcing - Margin Too Low") | (
                            merge["Product Status (Computed)"] == "Resourcing - Product Elevation In Progress") | (
                            merge["Product Status (Computed)"] == "Resourcing - Vendor Relation") | (
                            merge["Product Status (Computed)"] == "Resourcing") | (
                            merge["Item Type"] == "Powered Riding Toys") | (
                            merge["Reason"] == "Image might be considered to be RACY") | (
                            merge["Reason"] == "Field may contain suggestive and/or profane language.") | (
                            merge["Reason"] == "Illustrations/Logos are not acceptable images.") | (
                            merge["Reason"] == "Image may be a drawing.") | (
                            merge["Reason"] == "Image might be considered to be ADULT") | (
                            merge["Reason"] == "Image might be considered to be SPOOF") | (
                            merge["Reason"] == "Image might be considered to be MEDICAL") | (
                            merge["Reason"] == "Field may reference a weapon.") | (
                            merge["Reason"] == "Field may reference alcohol.") | (merge[
                                                                                      "Reason"] == "Field may indicate that the item does not comply with Target's inclusive merchandising policy") | (
                            merge["Reason"] == "This field is inherited from its parent and has errors") | (
                            merge["Reason"] == "This field is inherited from its parent and has errors") | (
                            merge["Reason"] == "The product was in a terminal state when the parent was versioned") | (
                            merge["Reason"] == "This item was put into suspended state")

        # creates condition which removes "Do Not Reorder" skus that are <= 130 qty
        cond2 = merge["Product Status (Computed)"].isin(
            ["Do Not Reorder – Exclude from Shopify", "Do Not Reorder – Safety Issue",
             "Do Not Reorder - Keep on ALL Marketplaces"]) & (merge["Total Quantity"] <= 130)

        # create dataframes containing all SKUs which meet conditions
        df3 = merge[cond1]
        df4 = merge[cond2]

        # concatenates all filtered data, removes duplicates
        df_all_rows = pd.concat([df3, df4]).drop_duplicates().reset_index(drop=True)

        # remove filtered SKUs from original dataset based on SKU + Reason
        unsorted = pd.merge(merge, df_all_rows, on=['Inventory Number', 'Reason'], how='left', indicator=True).query(
            "_merge != 'both'").drop('_merge', axis=1).drop(
            ['Product Title_y', 'Inventory_y', 'Data Update Status_y', 'Last Item Update_y', 'Error Category_y',
             'Partner Field Value_y', 'Submitted Field Name_y', 'Field Name_y', 'Partner Action_y',
             'Product Status (Computed)_y', 'Total Quantity_y', 'salsify:parent_id_y', 'Item Type_y'],
            axis=1).reset_index(drop=True)

        # sort by parent ID
        final = unsorted.sort_values(by=['Inventory Number'])

        final.to_excel("static/excel_files/downloads/output.xlsx")

    return render_template('public/templates/parser.html')

@app.route('/get-excel/<excel_download>')
def get_excel(excel_download):
    try:
        return send_from_directory(directory=config.CLIENT_EXCELS, filename=excel_download, as_attachment=False, path='/')
    except FileNotFoundError:
        abort(404)
    return "Ready for download"

if __name__ == '__main__':
    app.run()
