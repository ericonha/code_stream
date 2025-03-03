import os
from io import BytesIO
import openpyxl
import pandas as pd
import input_file
import numpy as np
import worker
import AP
import pdfkit
import streamlit as st
from datetime import datetime

def round_0_25(duration):
    duration = round_down_0_05(duration)

    if round(duration * 4) / 4 != duration:
        comparator = 0
        while comparator <= duration:
            comparator += 0.25
        return comparator
    return duration


def round_down_0_05(number):
    str_number = str(number)
    decimal_part = str_number.split(".")[-1]

    # Check if the number has more than one decimal place
    if len(decimal_part) > 1:
        return int(number * 20) / 20
    else:
        return number


def value_to_color(value):
    """Return a color based on the value. Near 0 is red, in the middle is blue, near 1 is green."""
    # Ensure value is between 0 and 1
    value = max(0, min(1, value))

    if value < 0.5:
        red = int((1 - value * 2) * 255)
        blue = int(value * 2 * 255)
        return f"rgb({red}, 0, {blue})"
    else:
        blue = int((1 - value) * 2 * 255)
        green = int((value - 0.5) * 2 * 255)
        return f"rgb(0, {green}, {blue})"


def format_euros(amount):
    return '€{:,.2f}'.format(amount)


def allocate_value(array, start_date, end_date, worker_id, value, years):
    # Convert start and end dates to datetime objects
    start_date = datetime.strptime(start_date, "%d.%m.%Y")
    end_date = datetime.strptime(end_date, "%d.%m.%Y")

    # Get the total number of days between start and end dates
    total_days = (end_date - start_date).days + 1

    # Loop over each year and calculate the value allocation
    for year in years:
        # Calculate the start and end dates of the current year
        year_int = int(year)
        year_start = datetime(year_int, 1, 1)
        year_end = datetime(year_int, 12, 31)

        # Determine the overlap of the year with the given start and end dates
        overlap_start = max(start_date, year_start)
        overlap_end = min(end_date, year_end)

        # Calculate the number of overlapping days in the year
        overlapping_days = (overlap_end - overlap_start).days + 1
        if overlapping_days > 0:
            # Calculate the portion of the value for this year
            year_allocation = (overlapping_days / total_days) * value
            # Assign the calculated value to the array
            year_index = int(int(year) - years[0])
            array[worker_id - 1][year_index] += year_allocation

    return array

# Add custom CSS for Streamlit to style the elements
st.markdown("""
    <style>
        .title {
            font-size: 24px;
            font-weight: bold;
            color: #333;
            text-align: center;
        }
        .button {
            background-color: #008CBA;
            color: white;
            padding: 10px 20px;
            border: none;
            cursor: pointer;
            border-radius: 3px;
            font-size: 14px;
        }
        .button:hover {
            background-color: #007B9A;
        }
        .input {
            font-size: 14px;
            padding: 10px;
            border-radius: 3px;
            border: 1px solid #ccc;
            margin-top: 10px;
        }
        .label {
            font-size: 14px;
            color: #333;
        }
    </style>
""", unsafe_allow_html=True)


# Functions to handle file selection and file paths
def upload_file(label):
    uploaded_file = st.file_uploader(label, type=["xlsx"])
    if uploaded_file is not None:
        return uploaded_file
    return None


def run_process(df, filepath, filepath_workers, name_of_output_file, entity):
    # create instance of AP to access functions
    ap1 = AP.AP()

    # get path to file PM and Workers
    input_file.get_arbeitspaket(df)
    input_file.get_all_names(df)

    # get in which colum each pms starts
    lista = input_file.get_dates(filepath)

    # get when each pm starts and ends, how many months per working year and in which year each pm starts and ends
    list_datas, lista_months, list_begin, list_end = input_file.get_dates_unix(df, lista)

    # input_file.get_color_of_company(df,filepath,entity)

    lista_datas_not_to_change = list_datas

    # save dates
    ap1.add_dates(list_datas[0], list_datas[1])

    # get which PM id this company will have to do
    ap1.get_hours(input_file.get_Company(df, entity))

    # get all the id PM titles names and take the first ad last one out (may refactor this later)
    ap1.Nr = input_file.get_arbeitspaket(df)
    ap1.Nr = ap1.Nr[1:]
    # ap1.Nr = ap1.Nr[0:-1]

    # get ids of pm
    ids = input_file.get_nrs(df)

    # clear workers info
    worker.list_of_workers.clear()

    # get workers hours
    input_file.get_workers_info(filepath_workers, lista_months)

    # sorte the worker from most expensive to least
    worker.sorte_workers()

    repeted_wh_ids = []

    # Check if intervals span across multiple years
    ap1.check_if_same_years(ids, repeted_wh_ids, list_begin, list_end)

    # get start and end year and save on ap1.year_start, ap1.year_end
    ap1.get_smallest_year()
    ap1.get_biggest_year()

    # cleaning dict for zettel infos
    AP.global_data_zettel_infos.clear()

    hours_year_work_every_one = []
    worker_hours = []
    for wk in worker.list_of_workers:
        worker_hours = []
        for i in range(0, ap1.year_end - ap1.year_start + 1):
            worker_hours.append(float(wk.hours_available[i].item()))
        hours_year_work_every_one.append(worker_hours)

    pre_define_workers = input_file.get_workers_pre_defined(df)

    h, ids_check, Nrs, pre_def = ap1.get_workers(lista_datas_not_to_change, ids, ap1.year_start, ap1.year_end, ap1.Nr,
                                                 entity,
                                                 df, pre_define_workers)

    # Generate HTML content with styling
    html_content = """
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body {
                    font-family: Arial, sans-serif;
                    margin: 20px;
                }
                h1 {
                    color: #333;
                    text-align: center;
                }
                table {
                    width: 100%;
                    border-collapse: collapse;
                    margin-bottom: 20px;
                }
                th, td {
                    border: 1px solid #ddd;
                    padding: 8px;
                    text-align: left;
                }
                th {
                    background-color: #f2f2f2;
                    color: #333;
                }
                tr:nth-child(even) {
                    background-color: #f9f9f9;
                }
                tr:hover {
                    background-color: #f5f5f5;
                }
            </style>
        </head>
        <body>
            <h1>Arbeits Packet Report</h1>
            <table>
                <tr>
                    <th>Id</th>
                    <th>AP</th>
                    <th>Start Date</th>
                    <th>End Date</th>
                    <th>Id Worker</th>
                    <th>WH</th>
                </tr>
        """

    print(sum(h))
    print(h)
    sum_test = 0
    ap_not_distribute = []
    years = np.linspace(ap1.year_start, ap1.year_end, ap1.year_end - ap1.year_start + 1)
    array_working_hours_per_year = np.zeros((len(worker.list_of_workers), len(years)))

    for w, wh, dates_st, dates_ft, Nr, id, pre_w in zip(ap1.workers, h, ap1.working_dates_start,
                                                        ap1.working_dates_end,
                                                        Nrs, ids_check, pre_def):
        if id in ids_check:
            w_id = 0
            if wh != 0:
                w_id = w.id
            row_color = "red" if w_id == 0 else "#ccffcc"

            if pre_w == 1:
                row_color = "blue"

            if w_id == 0:
                sum_test += wh
                ap_not_distribute.append(id)
            else:
                allocate_value(array_working_hours_per_year, dates_st, dates_ft, w_id, wh, years)
            html_content += f"""
                        <tr style="background-color: {row_color};">
                            <td>{id}</td>
                            <td>{Nr}</td>
                            <td>{dates_st}</td>
                            <td>{dates_ft}</td>
                            <td>{w_id}</td>
                            <td>{wh}</td>
                        </tr>
                """
    html_content += """
                </table>
                <div style="page-break-before: always;"></div>
            </body>
            </html>
            """
    print(sum_test)

    # Generate HTML content with styling for the second table
    html_content += """
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body {
                    font-family: Arial, sans-serif;
                }
                table {
                    width: 100%;
                    border-collapse: collapse;
                }
                th, td {
                    border: 1px solid #dddddd;
                    text-align: left;
                    padding: 10px;
                }
                th {
                    background-color: #f2f2f2;
                }
            </style>
        </head>
        <body>
            <h1>Sum Worker Report</h1>
            <table>
                <tr>
                    <th>year</th>
        """

    for i in range(len(worker.list_of_workers)):
        html_content += f"""
                        <th>Sum Worker {i + 1}</th>
            """
    html_content += f"""
                </tr>
        """

    p_p_y = np.array(lista_months) / 12

    sum_t = 0
    cost_project = 0

    new_array = []
    while len(worker.list_of_workers) != 0:
        lowest_index_elem = worker.Worker(1000, 0, 0)
        for element in worker.list_of_workers:
            if element.id < lowest_index_elem.id:
                lowest_index_elem = element
        new_array.append(lowest_index_elem)
        worker.list_of_workers.remove(lowest_index_elem)

    worker.list_of_workers = new_array

    # Add rows for sum worker data
    for i in range(len(years)):
        html_content += f"<tr>"
        html_content += f"<td>{int(years[i])}</td>"
        print(len(worker.list_of_workers))
        for j in range(len(worker.list_of_workers)):
            html_content += f"<td>{(hours_year_work_every_one[j][i])-(worker.list_of_workers[j].hours_available[i][0])}</td>"
            sum_t += hours_year_work_every_one[j][i] - worker.list_of_workers[j].hours_available[i][0]
            cost_project += float(hours_year_work_every_one[j][i] - worker.list_of_workers[j].hours_available[i][0]) * worker.list_of_workers[j].salary
        html_content += f"</tr>"

    # Add a row for total hours
    html_content += "<tr>"
    html_content += "<td><strong>Total</strong></td>"

    workers_total_hours = []

    for workers_t in array_working_hours_per_year:
        workers_total_hours.append(sum(workers_t))

    # Add totals for each worker
    for total in workers_total_hours:
        html_content += f"<td><strong>{total}</strong></td>"

    # Close the table and HTML body for the second table
    html_content += """
            </table>
            <table>
                <tr>
                   <th>Sum Total Hours</th>
                   <th>hours not distributed</th>
                   <th>APs not distributed</th>
                   <th>Cost of Project</th>
                   <th>Number of APs</th>
                </tr>
        """

    sum_t_b = sum_t
    print(sum_t_b)
    print(sum_t - sum_t_b)
    cost_project_formatted = format_euros(cost_project)
    aps_str = ""

    for aps in ap_not_distribute:
        aps_str += aps
        aps_str += ", "

    aps_str = aps_str[:-2]

    if aps_str == "":
        aps_str = "all aps distributed"

    html_content += f"""
            <tr>
                <td>{sum_t_b}</td>
                <td>{sum_test}</td>
                <td>{aps_str}</td>
                <td>{cost_project_formatted}</td>
                <td>{len(h)}</td>
            </tr>
        """

    html_content += """
        </table>
        </body>
        </html>
        """

    # Generate HTML content with styling for the second table
    html_content += """
            <html>
            <head>
                <meta charset="UTF-8">
                <style>
                    body {
                        font-family: Arial, sans-serif;
                    }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                    }
                    th, td {
                        border: 1px solid #dddddd;
                        text-align: left;
                        padding: 10px;
                    }
                    th {
                        background-color: #f2f2f2;
                    }
                </style>
            </head>
            <body>
                <h1>Dates distribution</h1>
                <table>
                    <tr>
                        <th>Worker Id</th>
                        <th>AP Id</th>
                        <th>Month</th>
                        <th>Year</th>
                        <th>Hours</th>
                        <th>PM (used to calculate hours 1 PM = 160 hours) </th>
                    </tr>
            """

    sorted_entries = sorted(
        AP.global_data_zettel_infos.items(),  # Sort by worker_id
        key=lambda x: (
            x[0],  # Sort by worker_id (x[0] is the worker_id)
            [entry['AP id'] for entry in x[1]],
            [entry['year'] for entry in x[1]],  # Sort by year (x[1] contains the entries for each worker)
            [entry['month'] for entry in x[1]]  # Sort by month (x[1] contains the entries for each worker)
        )
    )

    for worker_id, entries in sorted_entries:
        for entry in entries:
            # Extract month, hours, and PM
            month = entry['month']
            hours = entry['hours']
            year = entry['year']
            AP_id = entry['AP id']

            # Add a row for each entry
            html_content += f"""
            <tr>
                <td>{worker_id}</td>
                <td>{AP_id}</td>
                <td>{month}</td>
                <td>{year}</td>
                <td>{hours*160}</td>
                <td>{hours}</td>
            </tr>
            """

    # Save HTML content to a file
    with open("output.html", "w") as file:
        file.write(html_content)

    if len(name_of_output_file) == 0:
        print("Error name of pdf, it cannot be empty")
        exit(1)

    if len(name_of_output_file) > 100:
        print("Error name of pdf, it is way too big")
        exit(1)

    # Convert HTML to PDF
    file_name = name_of_output_file + "_" + entity + ".pdf"

    try:
        # Generate the PDF from the HTML content using pdfkit
        pdf_output = pdfkit.from_string(html_content)

        # Create a download button for the generated PDF
        st.download_button(
            label="Download PDF",
            data=pdf_output,
            file_name=file_name,
            mime="application/pdf"
        )
    except Exception as e:
        st.error(f"Error generating the PDF: {e}")


# Streamlit UI
st.title("Files Reader")

# File upload for AP and Worker files
ap_file = upload_file("Select Excel File for AP")
worker_file = upload_file("Select Excel File for Worker")

if ap_file is not None:
    st.write(f"Selected AP file: {ap_file.name}")
else:
    st.write("No AP file selected.")

if worker_file is not None:
    st.write(f"Selected Worker file: {worker_file.name}")
else:
    st.write("No Worker file selected.")

# When an AP file is uploaded, dynamically load entity names for dropdown
if ap_file:
    df_ap = input_file.get_file(ap_file)
    entities = input_file.get_all_names(df_ap)  # Fetch list of entities dynamically from the file
    entity = st.selectbox("Select Entity", entities)

# Input for output file name
output_name = st.text_input("Enter the name of the file to be saved")

import openpyxl
from io import BytesIO

# Run button for processing
if st.button("Run Process"):
    if ap_file and worker_file and entity != "Company/University/Hochschule" and output_name:
        if ap_file is not None:
            # Read the file into a pandas DataFrame
            try:
                # Read the .xlsx file directly from the uploaded file
                df = pd.read_excel(ap_file)
                run_process(df, ap_file, worker_file, output_name, entity)
            except Exception as e:
                st.error(f"Error processing the file: {e}")
    else:
        st.error("Please make sure all fields are filled out correctly.")





