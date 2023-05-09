from flask import Flask, render_template, request, redirect
import pandas as pd
import os

app = Flask(__name__)

if not os.path.isfile("demo.xlsx"):
    writer = pd.ExcelWriter("demo.xlsx", engine="xlsxwriter")
    df = pd.DataFrame({
        'Name': [],
        'Last Name': [],
        'Car Model': [],
        'Car Km': [],
        'Car Year': [],
        'Phone': []
    })
    df.to_excel(writer, index=False)
    writer.close()

@app.route("/api/car/save", methods=["POST", "GET"])
def save_car():
    first_name = request.form.get("fname")
    last_name = request.form.get("lname")
    car_model = request.form.get("cmodel")
    car_km = request.form.get("ckm")
    car_year = request.form.get("cyear")
    phone_number = request.form.get("phonnum")


    reader = pd.read_excel('demo.xlsx', engine="openpyxl")
    df = pd.DataFrame({
        'Name': [first_name],
        'Last Name': [last_name],
        'Car Model': [car_model],
        'Car Km': [car_km],
        'Car Year': [car_year],
        'Phone': [phone_number]})
    writer = pd.ExcelWriter("demo.xlsx", engine="openpyxl", mode="a", if_sheet_exists="overlay")
    df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=len(reader) + 1)
    writer.close()

    return redirect("/")


@app.route("/")
def index():
    return render_template("index.html")


if __name__ == "__main__":
    app.run()
