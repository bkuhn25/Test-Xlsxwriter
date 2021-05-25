from flask import Flask, render_template, redirect, url_for, flash, abort, request
from flask_bootstrap import Bootstrap
import os
import xlsxwriter

app = Flask(__name__)


@app.route('/', methods=["GET", "POST"])
def home():
    workbook = xlsxwriter.Workbook("test1.xlsx")
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, 25)
    workbook.close()
    return render_template("test.html")


if __name__ == "__main__":
    app.run()
