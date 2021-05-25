from flask import Flask, render_template, redirect, url_for, flash, abort, request, send_from_directory
from flask_bootstrap import Bootstrap
import os
from datetime import datetime
from os import listdir
from os.path import isfile, join
import xlsxwriter

app = Flask(__name__)


def create_directories():
    """Adds a directory for the day the program is being used to store all of the combined schedules
        hat are created"""

    current_directory = os.getcwd()
    combined_schedules_dir = os.path.join(current_directory, "static", "Files")

    today_date = datetime.now().strftime("%I-%M-%S")
    today_schedules = os.path.join(combined_schedules_dir, today_date)

    if not os.path.exists(combined_schedules_dir):
        os.makedirs(combined_schedules_dir)

    return today_schedules


def directory_path():
    current_directory = os.getcwd()
    combined_schedules_dir = os.path.join(current_directory, "static", "Files")

    return combined_schedules_dir


@app.route('/', methods=["GET", "POST"])
def home():
    workbook = xlsxwriter.Workbook(f"{create_directories()}.xlsx")
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, 25)
    workbook.close()
    return render_template("index.html")


@app.route('/check-files', methods=["GET"])
def check_files():
    onlyfiles = [f for f in listdir(directory_path()) if isfile(join(directory_path(), f))]
    print(onlyfiles)
    return render_template("excel_files.html", files=onlyfiles)


@app.route('/download', methods=["GET"])
def download():
    file = request.args.get("filename")
    return send_from_directory("static", filename=f"Files/{file}")


if __name__ == "__main__":
    app.run()
