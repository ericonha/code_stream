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
import shutil

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

def get_german_month(english_month):
    months = {
        "January": "Januar",
        "February": "Februar",
        "March": "März",
        "April": "April",
        "May": "Mai",
        "June": "Juni",
        "July": "Juli",
        "August": "August",
        "September": "September",
        "October": "Oktober",
        "November": "November",
        "December": "Dezember"
    }
    return months.get(english_month, "Invalid month")


month_map = {
    'January': 1, 'February': 2, 'March': 3, 'April': 4,
    'May': 5, 'June': 6, 'July': 7, 'August': 8,
    'September': 9, 'October': 10, 'November': 11, 'December': 12
}


# Só para ordenação: converte '3.1' → (3, 1), mas sem modificar o valor original
def ap_id_sort_key(ap_id_str):
    return tuple(map(int, ap_id_str.split('.')))


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

    New_Nrs, New_ids, New_year_start, New_year_end, New_pre_define_workers, New_hours, Shuffle_to_Original_Index = AP.shuffle_aligned_lists(ap1.Nr, ids, lista_datas_not_to_change[0], lista_datas_not_to_change[1], pre_define_workers, ap1.hours)

    h, ids_check, Nrs, pre_def = ap1.get_workers([New_year_start,New_year_end], New_ids, ap1.year_start, ap1.year_end, New_Nrs,
                                                 entity,
                                                 df, New_pre_define_workers, New_hours)

    #h, ids_check, Nrs, pre_def = ap1.get_workers(lista_datas_not_to_change, ids, ap1.year_start, ap1.year_end, ap1.Nr,
    #                                             entity,
    #                                             df, pre_define_workers, ap1.hours)

    # For example, reconstruct data arrays
    order_aps_final = sorted(AP.order_aps, key=lambda x: tuple(map(float, x[1].split("."))))

    restored_Nrs = [x[0] for x in order_aps_final]
    restored_ids = [x[1] for x in order_aps_final]
    restored_start = [x[2] for x in order_aps_final]
    restored_end = [x[3] for x in order_aps_final]
    restored_dates = [x[4] for x in order_aps_final]
    restored_hours = [x[5] for x in order_aps_final]
    restored_workers = [x[6] for x in order_aps_final]




    html_content_1 = """
    <html lang="de">
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
                table-layout: fixed;
            }
            th, td {
                border: 1px solid #ddd;
                padding: 8px;
                text-align: left;
                overflow: hidden;
                white-space: normal;
                word-break: normal;      /* Only break at natural points */
                hyphens: auto;           /* Allow hyphenation */
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
            td:nth-child(2) {
                font-size: 12px;  /* Adjust the font size as needed */
            }
        </style>
    </head>
    <body>
        <h1>Arbeitspaketbericht</h1>
        <table>
            <colgroup>
                <col style="width:10%">    <!-- Id -->
                <col style="width:37%">   <!-- AP (WIDER) -->
                <col style="width:17%">   <!-- Startdatum -->
                <col style="width:17%">   <!-- Enddatum -->
                <col style="width:10%">    <!-- Id Arbeiter -->
                <col style="width:8%">    <!-- WH -->
            </colgroup>
            <tr>
                <th>Id</th>
                <th>AP</th>
                <th>Startdatum</th>
                <th>Enddatum</th>
                <th>Id Arbeiter</th>
                <th>WH</th>
            </tr>
    """

    print(sum(h))
    print(len(h))
    print(len(ap1.working_dates_start))
    print(h)
    sum_test = 0
    sum_test2 = 0
    ap_not_distribute = []
    years = np.linspace(ap1.year_start, ap1.year_end, ap1.year_end - ap1.year_start + 1)
    array_working_hours_per_year = np.zeros((len(worker.list_of_workers), len(years)))

    for w, wh, dates_st, dates_ft, Nr, id in zip(restored_workers,restored_hours,restored_start,restored_end,
                                                 restored_Nrs,restored_ids):
        html_content_1 += f"""
                                <tr style="background-color: {"#ccffcc"};">
                                    <td>{id}</td>
                                    <td>{Nr}</td>
                                    <td>{dates_st}</td>
                                    <td>{dates_ft}</td>
                                    <td>{w.id}</td>
                                    <td>{round(wh,2)}</td>
                                </tr>
                        """
    html_content_1 += """
                    </table>
                    <div style="page-break-before: always;"></div>
                </body>
                """

    # Generate HTML content with styling for the second table
    html_content_1 += """
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
                    font-size: 12px;
                }
                th {
                    background-color: #f2f2f2;
                }
            </style>
        </head>
        <body>
            <h1>Summen arbeiterbericht</h1>
            <table>
                <tr>
                    <th>Jahr</th>
        """

    for i in range(len(worker.list_of_workers)):
        html_content_1 += f"""
                        <th>Summen arbeiter {i + 1}</th>
            """
    html_content_1 += f"""
                </tr>
        """

    p_p_y = np.array(lista_months) / 12

    sum_t = 0
    cost_project = 0

    new_array = []
    new_array_hours = []
    while len(worker.list_of_workers) != 0:
        lowest_index_elem = worker.Worker(1000, 0, 0, 0, "", "",0)
        for element in worker.list_of_workers:
            if element.id < lowest_index_elem.id:
                lowest_index_elem = element
        new_array.append(lowest_index_elem)
        index = worker.list_of_workers.index(lowest_index_elem)
        new_array_hours.append(hours_year_work_every_one[index])

        worker.list_of_workers.remove(lowest_index_elem)
        hours_year_work_every_one.remove(hours_year_work_every_one[index])

    worker.list_of_workers = new_array
    hours_year_work_every_one = new_array_hours

    # Add rows for sum worker data
    for i in range(len(years)):
        html_content_1 += f"<tr>"
        html_content_1 += f"<td>{int(years[i])}</td>"
        for j in range(len(worker.list_of_workers)):
            html_content_1 += f"<td>{round((hours_year_work_every_one[j][i]) - (worker.list_of_workers[j].hours_available[i][0]),2)}</td>"
            sum_t += hours_year_work_every_one[j][i] - worker.list_of_workers[j].hours_available[i][0]
            cost_project += round(float(hours_year_work_every_one[j][i] - worker.list_of_workers[j].hours_available[i][0]) * \
                            worker.list_of_workers[j].salary,2)
        html_content_1 += f"</tr>"

    # Add a row for total hours
    html_content_1 += "<tr>"
    html_content_1 += "<td><strong>Total</strong></td>"

    workers_total_hours = []
    index_year = 0

    for workers_t in worker.list_of_workers:
        w_hours = []
        for index_ele in range(len(workers_t.hours_available)):
            w_hours.append(workers_t.hours_available[index_ele][0])
        workers_total_hours.append(w_hours)

    total_w = []
    for index_ele in range(len(workers_total_hours)):
        total_w.append(sum(np.array(hours_year_work_every_one[index_ele]) - np.array(workers_total_hours[index_ele])))



    # Add totals for each worker
    for total in total_w:
        html_content_1 += f"<td><strong>{round(total,2)}</strong></td>"

    # Close the table and HTML body for the second table
    html_content_1 += """
            </table>
            <table>
                <tr>
                   <th>Summe der Gesamtstunden</th>
                   <th>Stunden nicht verteilt</th>
                   <th>APs nicht verteilt</th>
                   <th>Projektkosten</th>
                   <th>Anzahl der APs</th>
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
        aps_str = "Alle APs verteilt"

    html_content_1 += f"""
            <tr>
                <td>{round(sum_t_b,2)}</td>
                <td>{sum_test}</td>
                <td>{aps_str}</td>
                <td>{cost_project_formatted}</td>
                <td>{len(h)}</td>
            </tr>
        """

    html_content_1 += """
        </table>
        </body>
        </html>
        """

    # Generate HTML content with styling for the second table
    html_content_2 = ""
    html_content_2 += """
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
                <h1>Terminverteilung</h1>
                <table>
                    <tr>
                        <th>Arbeiter</th>
                        <th>AP-Id</th>
                        <th>Monat</th>
                        <th>Jahr</th>
                        <th>Stunden</th>
                        <th>PM (wird zur Stundenberechnung verwendet: 1 PM = 160 Stunden) </th>
                    </tr>
            """

    #sorted_entries = sorted(
    #    AP.global_data_zettel_infos.items(),  # Sort by worker_id
    #    key=lambda x: (
    #        x[0],  # Sort by worker_id (x[0] is the worker_id)
    #        [entry['AP id'] for entry in x[1]],
    #        [entry['year'] for entry in x[1]],  # Sort by year (x[1] contains the entries for each worker)
    #        [entry['month'] for entry in x[1]]  # Sort by month (x[1] contains the entries for each worker)
    #    )
    #)

    all_entries = [
        entry
        for worker_entries in AP.global_data_zettel_infos.values()
        for entry in worker_entries
    ]

    sorted_entries = sorted(
        all_entries,
        key=lambda e: (
            e['worker_id'],
            ap_id_sort_key(e['AP id']),  # usa parse para ordenar, mas não altera o valor
            month_map.get(e['month'], 0),  # mês como número
            e['year'],
            e['hours']
        )
    )




    for entry in sorted_entries:
        # Extract month, hours, and PM
        month = entry['month']
        hours = entry['hours']
        year = entry['year']
        AP_id = entry['AP id']
        worker_id = entry['worker_id']

        if hours == 0:
            continue

        name = ""

        for wks in worker.list_of_workers:
            if wks.id == worker_id:
                name = str(wks.name) + " " + str(wks.surname)

        # Add a row for each entry
        html_content_2 += f"""
         <tr>
            <td>{name}</td>
            <td>{AP_id}</td>
            <td>{get_german_month(month)}</td>
            <td>{year}</td>
            <td>{round(hours,2) * 160}</td>
            <td>{round(hours,2)}</td>
        </tr>
        """
    html_content_2 += """
             </table>
         </body>
         </html>
         """

    # Generate HTML content with styling for the second table
    html_content_2 += """
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
            <h1>Monatlicher Arbeiterbericht</h1>
            <table>
                <tr>
                    <th>Arbeiter</th>
                    <th>Stunden</th>
                    <th>Jahr</th>
                    <th>Monat</th>
                 </tr>
        """

    months_german = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober",
                     "November", "Dezember"]

    last_month = 12
    for wk in worker.list_of_workers:
        first_year = 0
        for year_idx, year in enumerate(years):
            if first_year == 0:
                months_to_iterate = 12 - lista_months[0]
                first_year = 1
                last_month = 12
            else:
                months_to_iterate = 0
                last_month = lista_months[year_idx]

            for i in range(months_to_iterate, last_month):
                month_idx = i
                hours = 1 - wk.hours_available_per_month[year_idx][month_idx]

                html_content_2 += f"""
                <tr>
                    <td>{str(wk.name) + " " + str(wk.surname)}</td>
                    <td>{round(hours,2)}</td>
                    <td>{int(year)}</td>
                    <td>{months_german[month_idx]}</td>
                </tr>
                """

    html_content_2 += """
            </table>
        </body>
        </html>
    """

    #path library
    wkhtmltopdf_path = shutil.which("wkhtmltopdf")
    
    if wkhtmltopdf_path is None:
        raise RuntimeError("wkhtmltopdf not found")
    else:
        print(wkhtmltopdf_path)

    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)

    # Save HTML content to a file
    with open("output.html", "w") as file:
        file.write(html_content_1)

    with open("output.html", "w") as file:
        file.write(html_content_2)

    if len(name_of_output_file) == 0:
        print("Error name of pdf, it cannot be empty")
        exit(1)

    if len(name_of_output_file) > 100:
        print("Error name of pdf, it is way too big")
        exit(1)

    # Convert HTML to PDF
    file_name_1 = name_of_output_file + "_datum" + "_" + entity + ".pdf"
    file_name_2 = name_of_output_file + "_organizer" + "_" + entity + ".pdf"

    try:
        # Generate the PDF from the HTML content using pdfkit
        pdf_output_1 = pdfkit.from_string(html_content_1,False, configuration=config)
        pdf_output_2 = pdfkit.from_string(html_content_2,False, configuration=config)

        # Create a download button for the generated PDF
        st.download_button(
            label=file_name_1,
            data=pdf_output_1,
            file_name=file_name_1,
            mime="application/pdf"
        )

        st.download_button(
            label=file_name_2,
            data=pdf_output_2,
            file_name=file_name_2,
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
    entity = st.selectbox("Select Entity", entities[0:-1])

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





