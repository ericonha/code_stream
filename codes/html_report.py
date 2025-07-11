from weasyprint import HTML
from typing import List
from worker import WorkPackage, Worker

def generate_full_html_report(aps_list: List[WorkPackage], workers_list: List[Worker], start_year: int):
    html = """
    <html lang="de">
    <head>
        <meta charset="UTF-8">
        <style>
            body {
                font-family: Arial, sans-serif;
                margin: 20px;
                font-size: 13px;
            }
            h1, h2 {
                color: #333;
                text-align: center;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 30px;
                table-layout: fixed;
            }
            th, td {
                border: 1px solid #ddd;
                padding: 6px;
                text-align: center;
                word-break: break-word;
            }
            th {
                background-color: #f2f2f2;
                color: #333;
            }
            tr:nth-child(even) {
                background-color: #f9f9f9;
            }
        </style>
    </head>
    <body>
        <h1>Arbeitspaketbericht</h1>
        <table>
            <thead>
                <tr>
                    <th>AP Id</th>
                    <th>AP Titel</th>
                    <th>Startdatum</th>
                    <th>Enddatum</th>
                    <th>Arbeiter-ID</th>
                    <th>Arbeiter-Name</th>
                    <th>WH</th>
                </tr>
            </thead>
            <tbody>
    """

    total_unassigned_wh = 0.0
    aps_with_unassigned = []

    for ap in aps_list:
        ids = ap.assigned_worker_ids if hasattr(ap, "assigned_worker_ids") else []
        names = [n.strip() for n in (ap.assigned_worker_name or "").split(",")]
        total_ids = len(ids)
        if total_ids > 1:
            for i, wid in enumerate(ids):
                name = names[i] if i < len(names) else "?"
                partial_pm = ap.required_pm / total_ids
                partial_pm = workers_list[ap.assigned_worker_ids[i]-1].assignment_log_hours_per_ap[ap.id]
                row_color = "#ffe6e6" if wid == 0 else "#ccffcc"
                if wid == 0:
                    total_unassigned_wh += partial_pm
                    aps_with_unassigned.append(ap)
                html += f"""
                <tr style="background-color:{row_color};">
                    <td>{ap.id}</td>
                    <td>{ap.title[:11] + "..."}</td>
                    <td>{ap.start_date}</td>
                    <td>{ap.end_date}</td>
                    <td>{wid}</td>
                    <td>{name}</td>
                    <td>{round(partial_pm,3):.2f}</td>
                </tr>
                """
        else:
            wid = ids[0] if ids else "?"
            name = names[0] if names else "?"
            row_color = "#ffe6e6" if wid == 0 else "#ccffcc"
            if wid == 0:
                total_unassigned_wh += ap.required_pm
                aps_with_unassigned.append(ap)
            html += f"""
            <tr style="background-color:{row_color};">
                <td>{ap.id}</td>
                <td>{ap.title[:11] + "..."}</td>
                <td>{ap.start_date}</td>
                <td>{ap.end_date}</td>
                <td>{wid}</td>
                <td>{name}</td>
                <td>{ap.required_pm:.2f}</td>
            </tr>
            """

    html += """
            </tbody>
        </table>

        <div style="page-break-before: always;"></div>

        <h2>Mitarbeiterstatistik</h2>
        <table>
            <thead>
                <tr>
                    <th>Arbeiter-ID</th>
                    <th>Name</th>
    """

    num_years = workers_list[0].years if workers_list else 0
    for y in range(num_years):
        html += f"<th>{start_year + y}</th>"
    html += "<th>Summe</th></tr></thead><tbody>"

    total_project_cost = 0
    total_hours = 0

    for w in workers_list:
        html += f"<tr><td>{w.id}</td><td>{w.name} {w.surname}</td>"
        for hours in w.summed_hours_worked_per_year:
            html += f"<td>{hours:.2f}</td>"
        html += f"<td><b>{w.summed_hours_worked_total:.2f}</b></td></tr>"
        total_hours += w.summed_hours_worked_total
        total_project_cost += w.summed_hours_worked_total * w.salary

    html += f"""
        <tr style="font-weight: bold; background-color: #e8e8e8;">
            <td colspan="2">Gesamt</td>
    """
    for y in range(num_years):
        total_per_year = sum(w.summed_hours_worked_per_year[y] for w in workers_list)
        html += f"<td>{total_per_year:.2f}</td>"
    html += f"<td>{total_hours:.3f}</td></tr>"

    html += """
            </tbody>
        </table>
    """

    html += f"""
        <h2>Projektübersicht</h2>
        <table>
            <thead>
                <tr>
                    <th>Projektkosten (€)</th>
                    <th>Anzahl APs</th>
                    <th>Gesamt WH</th>
                    <th style="color:red;">WH Nicht Zugewiesen</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>{total_project_cost:.2f}</td>
                    <td>{len(aps_list)}</td>
                    <td>{total_hours:.2f}</td>
                    <td style="color:red;">{total_unassigned_wh:.2f}</td>
                </tr>
            </tbody>
        </table>
    """

    if total_unassigned_wh > 0:
        html += "<h2 style='color:red;'>Nicht vollständig verteilte APs</h2><ul>"
        for ap in aps_with_unassigned:
            html += f"<li>AP {ap.id} – {ap.title}</li>"
        html += "</ul>"

    html += "</body></html>"
    return html

def save_pdf_report(aps_list: List[WorkPackage], workers: List[Worker], output_path: str = "arbeitspaket_bericht.pdf", start_year: int = 0):
    html_content = generate_full_html_report(aps_list, workers, start_year)
    HTML(string=html_content).write_pdf(output_path)
    print(f"✅ PDF saved to: {output_path}")


from weasyprint import HTML
from typing import List
from worker import Worker
import re

def ap_id_sort_key(ap_id: str):
    """Sort AP IDs like '1.1', '1.10', etc., in natural numeric order."""
    return tuple(int(part) if part.isdigit() else part for part in re.split(r'[^\d]+', ap_id) if part != '')


def generate_detailed_worker_report_html(workers: List[Worker], start_year: int) -> str:
    month_names = [
        "Januar", "Februar", "März", "April", "Mai", "Juni",
        "Juli", "August", "September", "Oktober", "November", "Dezember"
    ]

    html = """
    <html lang="de">
    <head>
        <meta charset="UTF-8">
        <style>
            body { font-family: Arial, sans-serif; margin: 20px; font-size: 13px; }
            h1, h2 { color: #333; text-align: center; }
            table { width: 100%; border-collapse: collapse; margin-bottom: 30px; table-layout: fixed; }
            th, td { border: 1px solid #ddd; padding: 6px; text-align: center; word-break: break-word; }
            th { background-color: #f2f2f2; color: #333; }
            tr:nth-child(even) { background-color: #f9f9f9; }
            .page-break { page-break-before: always; }
        </style>
    </head>
    <body>
        <h1>Detaillierter Arbeitsbericht</h1>
    """

    # TABLE 1: Worker, AP-ID, Month, Year, Hours, PM
    rows_ap_view = []
    for worker in workers:
        for (year_idx, month_idx), ap_data in worker.assignment_log.items():
            for ap_id, pm in ap_data:
                hours = pm * 160
                rows_ap_view.append((
                    worker.name + " " + worker.surname,
                    ap_id,
                    month_names[month_idx],
                    start_year + year_idx,
                    hours,
                    pm
                ))

    # ✅ Sort by AP-ID using natural sort
    rows_ap_view.sort(key=lambda x: ap_id_sort_key(x[1]))

    html += """
    <h2>Beteiligung pro Arbeitspaket</h2>
    <table>
        <thead>
            <tr>
                <th>Arbeiter</th>
                <th>AP-Id</th>
                <th>Monat</th>
                <th>Jahr</th>
                <th>Stunden</th>
                <th>PM</th>
            </tr>
        </thead>
        <tbody>
    """

    for row in rows_ap_view:
        html += f"""
        <tr>
            <td>{row[0]}</td>
            <td>{row[1]}</td>
            <td>{row[2]}</td>
            <td>{row[3]}</td>
            <td>{row[4]:.2f}</td>
            <td>{row[5]:.2f}</td>
        </tr>
        """

    html += "</tbody></table>"

    # TABLE 2: Per-worker summary per month
    for worker in workers:
        html += '<div class="page-break"></div>'
        html += f"<h2>Zusammenfassung – {worker.name} {worker.surname}</h2>"
        html += """
        <table>
            <thead>
                <tr>
                    <th>Jahr</th>
                    <th>Monat</th>
                    <th>Stunden</th>
                </tr>
            </thead>
            <tbody>
        """

        monthly_totals = {}
        for (year_idx, month_idx), ap_data in worker.assignment_log.items():
            year = start_year + year_idx
            month = month_names[month_idx]
            hours = sum(pm * 160 for _, pm in ap_data)
            key = (year, month)
            monthly_totals[key] = monthly_totals.get(key, 0) + hours

        for (year, month), hours in sorted(monthly_totals.items()):
            html += f"""
            <tr>
                <td>{year}</td>
                <td>{month}</td>
                <td>{hours:.2f}</td>
            </tr>
            """

        html += "</tbody></table>"

    html += "</body></html>"
    return html


def save_worker_assignment_pdf(workers: List[Worker], output_path: str = "mitarbeiter_bericht.pdf", start_year: int = 0):
    html_content = generate_detailed_worker_report_html(workers, start_year=start_year)
    HTML(string=html_content).write_pdf(output_path)
    print(f"✅ Mitarbeiter-PDF gespeichert unter: {output_path}")



