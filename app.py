from flask import Flask, render_template, request, redirect, send_from_directory, abort
import os
from werkzeug.utils import secure_filename
import config
import pandas as pd
import openpyxl

app = Flask(__name__)


# Specifies allowed filetypes (see environment vars)
def allowed_excel(filename):
    if not "." in filename:
        return False
    ext = filename.rsplit(".", 1)[1]
    if ext.upper() in config.ALLOWED_EXTENSIONS:
        return True
    else:
        return False


# Handler for excel file uploads or whatever else
@app.route('/', methods=["GET", "POST"])
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


@app.route('/get-orders', methods=['POST', 'GET'])
# Refreshes API key and stores in env variable. Need to link to a button
def refreshToken():
    url = "https://api.channeladvisor.com/oauth2/token"

    auth_str = '{}:{}'.format(config.app_id, config.shared_secret)
    b64_auth_str = base64.urlsafe_b64encode(auth_str.encode()).decode()

    payload = 'grant_type=refresh_token&refresh_token=' + config.refresh_token
    headers = {
        'Authorization': 'Basic ' + b64_auth_str,
        'Content-Type': 'application/x-www-form-urlencoded'
    }

    response = requests.request("POST", url, headers=headers, data=payload)

    print(response.text)
    token_json = json.loads(response.text)
    print(token_json['access_token'])

    f = open("config2.py", "w")
    f.write("refreshed_token = \"" + token_json['access_token'] + "\"")
    f.close()


# Main script. Gets unshipped orders from ChAd API. Output should be displayed on the page
def getOrders():
    if request.method == "GET":
        r = requests.get(
            "https://api.channeladvisor.com/v1/Orders?$filter=ShippingStatus eq 'Unshipped' and ProfileID eq 32001378&access_token=" + config2.refreshed_token)
        list_of_attributes = r.text
        attributes = json.loads(list_of_attributes)

        # Converts API response to JSON
        with open("data.json", "w") as write:
            json.dump(attributes, write)

        # Defines list variables
        full_list = attributes["value"]
        df = pd.DataFrame(full_list)
        order_ids = df['ID']
        order_data = []
        sku_list = []

        # Filters response and iterates through lists to retrieve each SKU
        def retrieveOrderItems():
            for x in order_ids:
                order_items = requests.get("https://api.channeladvisor.com/v1/Orders(" + str(
                    x) + ")/Items?$filter=ProfileID eq 32001378&access_token=" + config2.refreshed_token)
                order_items_json = json.loads(order_items.text)
                order_data.append(order_items_json["value"])
            for list in order_data:
                for sku in list:
                    sku_list.append(sku)
            df2 = pd.DataFrame(sku_list)
            return df2

        # Calls function and creates dataframe from returned data. Drops unnecessary columns
        unshipped_skus = retrieveOrderItems()
        unshipped_skus.drop(
            ['ProductID', 'SiteOrderItemID', 'SellerOrderItemID', 'UnitPrice', 'TaxPrice', 'ShippingPrice',
             'ShippingTaxPrice', 'RecyclingFee', 'UnitEstimatedShippingCost', 'GiftMessage', 'GiftNotes', 'GiftPrice',
             'GiftTaxPrice', 'IsBundle', 'ItemURL', 'HarmonizedCode'], axis=1, inplace=True)
        unshipped_skus.to_excel("ID data.xlsx", sheet_name="Sheet1")

        # Loads entire ChannelAdvisor inventory and merges based on SKU. Outputs to Activewear Upload.xlsx
        chad_inv = pd.read_excel('chad_inv.xlsx')
        activewear_skus = pd.merge(unshipped_skus, chad_inv, how='left', on='Sku')
        print(activewear_skus)
        activewear_skus.to_excel("Activewear Upload.xlsx", sheet_name="Sheet1")

    return render_template('public/templates/get-orders.html')


# Converts orders to JSON which SSActivewear can read
def ConvertOrders():
    df = pd.read_excel('Activewear Upload.xlsx')

    # Drops all rows that have no ActiveWear SKU and extraneous columns
    df2 = df.dropna(subset=['Attribute1Value'])
    df2.drop(
        ['ID', 'ProfileID', 'OrderID', 'SiteListingID', 'Title', 'Classification', 'Attribute1Name', 'Unnamed: 0'],
        axis=1, inplace=True)
    df2.reset_index(drop=True, inplace=True)

    # with pd.option_context('display.max_rows', None, 'display.max_columns', None):  # more options can be specified also
    #     print(df2)

    # creates JSON object to send to SSActivewear. Does a lot of reformatting of the DataFrame object and converts to JSON
    df2.drop(['Sku', 'Attribute2Value', 'Attribute2Name'], axis=1, inplace=True)
    df2.rename(columns={'Quantity': 'qty', 'Attribute1Value': 'identifier'}, inplace=True)
    df2 = df2[['identifier', 'qty']]
    order = df2.to_json('orders.json', orient="records")


# Submits final order to SSActivewear
def submitOrder():
    f = open('orders.json')
    lines = json.load(f)

    url = "https://api.ssactivewear.com/v2/orders/"

    payload = json.dumps({
        "shippingAddress": {
            "customer": "Company ABC",
            "attn": "John Doe",
            "address": "123 Main St",
            "city": "Bolingbrook",
            "state": "IL",
            "zip": "60440",
            "residential": True
        },
        "shippingMethod": "1",
        "shipBlind": False,
        "poNumber": "Online Test",
        "RejectLineErrors": "false",
        "emailConfirmation": "aforkbends@gmail.com",
        "rejectLineErrors_Email": "true",
        "testOrder": True,
        "autoselectWarehouse": True,
        "lines": lines
    })
    headers = {
        'Authorization': 'Basic ' + config.ssa_api_key,
        'Content-Type': 'application/json',
        'Cookie': '__cf_bm=_BOX3o_owW3dCpquJ8apGRYK0MUxc9DXLDbylhj55qo-1655413418-0-AVqP+NDYBun21+X5wUJWXech2e6q/YYoKCVhAhozg1Zrq93i49AF0Vu+DzOOjzvYBnadIyN0he92Ob3MVROG/bnVYnFkUxX2aqZiQYSGdBBw'
    }

    response = requests.request("POST", url, headers=headers, data=payload)

    print(response.text)


# Handler for downloading excel files or whatever else
@app.route('/get-excel/<excel_download>')
def get_excel(excel_download):
    try:
        return send_from_directory(directory=config.CLIENT_EXCELS, filename=excel_download, as_attachment=False,
                                   path='/')
    except FileNotFoundError:
        abort(404)
    return "Ready for download"


if __name__ == '__main__':
    app.run()
