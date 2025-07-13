import streamlit as st
import pandas as pd
import copy
import zipfile
import random
import io

import worker
import assignments
import html_report


def company_in_file(uploaded_file, company: str) -> bool:
    if uploaded_file is None or not company.strip():
        return False
    try:
        df = pd.read_excel(uploaded_file, header=None)
        if df.shape[0] <= 3:
            st.warning("🚫 Die Datei hat weniger als 4 Zeilen – Headerzeile fehlt?")
            return False

        header_row = df.iloc[3].astype(str).str.strip()
        return any(company.lower() in col.lower() for col in header_row if isinstance(col, str))
    except Exception as e:
        st.error(f"❌ Fehler beim Verarbeiten der Datei: {e}")
        return False


def main():
    st.set_page_config(page_title="Arbeitsplan Optimierer", layout="centered")
    st.title("🛠️ Arbeitsplan Optimierer")

    with st.sidebar:
        st.header("⚙️ Einstellungen")
        company_name = st.text_input("Firmenname", value="")
        rounds = st.number_input("Anzahl der Optimierungsdurchläufe", min_value=100, value=100)

    ap_file = st.file_uploader("📄 Arbeitsplan hochladen (Excel)", type=["xlsx"])
    worker_file = st.file_uploader("👷 Personalkosten hochladen (Excel)", type=["xlsx"])

    valid_inputs = (
        company_name.strip() != ""
        and rounds > 0
        and ap_file is not None
        and worker_file is not None
        and company_in_file(ap_file, company_name)
    )

    if ap_file and company_name.strip() != "" and not company_in_file(ap_file, company_name):
        st.warning("⚠️ Firmenname nicht in der Datei gefunden. Bitte überprüfen.")

    if valid_inputs:
        if st.button("🚀 Optimierung starten"):
            ap_bytes = ap_file.read()
            ap_buffer = io.BytesIO(ap_bytes)

            df_ap = pd.read_excel(ap_file, header=None)
            df_workers = pd.read_excel(worker_file)

            aps_list_orig = worker.extract_work_packages_from_dataframe(df_ap, company=company_name, file_buffer=ap_file)
            months = worker.month_per_year(df_ap)
            workers_list_orig = worker.extract_workers_from_dataframe(df_workers, months)
            start_year = int(aps_list_orig[0].start_date[6:])

            best_aps_list = None
            best_workers_list = None
            best_cost = -1
            successful_runs = 0
            found = False

            with st.spinner("⏳ Optimierung läuft..."):
                for _ in range(rounds):
                    aps_list = copy.deepcopy(aps_list_orig)
                    workers_list = copy.deepcopy(workers_list_orig)

                    original_order = {ap.id: idx for idx, ap in enumerate(aps_list)}
                    random.shuffle(aps_list)

                    for ap in aps_list:
                        assignments.assign_ap_to_workers(ap, workers_list, start_year)

                    aps_list.sort(key=lambda ap: original_order.get(ap.id, 9999))
                    all_distributed = all("Nicht zugewiesen" not in (ap.assigned_worker_name or "") for ap in aps_list)
                    total_cost = sum(w.summed_hours_worked_total * w.salary for w in workers_list)

                    if all_distributed:
                        successful_runs += 1
                        if total_cost > best_cost:
                            best_cost = total_cost
                            best_aps_list = copy.deepcopy(aps_list)
                            best_workers_list = copy.deepcopy(workers_list)
                            found = True
                    elif total_cost > best_cost:
                        best_cost = total_cost
                        best_aps_list = copy.deepcopy(aps_list)
                        best_workers_list = copy.deepcopy(workers_list)

            if best_aps_list:
                st.success(f"✅ Beste Lösung erfolgreich gefunden")

                # Create in-memory PDF buffers
                ap_pdf_buffer = io.BytesIO()
                worker_pdf_buffer = io.BytesIO()

                html_report.save_pdf_report(best_aps_list, best_workers_list, ap_pdf_buffer, start_year=start_year)
                html_report.save_worker_assignment_pdf(best_workers_list, worker_pdf_buffer, start_year=start_year)

                # ZIP everything in-memory
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zipf:
                    zipf.writestr("arbeitspaket_bericht.pdf", ap_pdf_buffer.getvalue())
                    zipf.writestr("mitarbeiter_report.pdf", worker_pdf_buffer.getvalue())

                zip_buffer.seek(0)
                st.download_button("📦 ZIP herunterladen", zip_buffer, file_name="berichte.zip", mime="application/zip")

                if not found:
                    st.warning("⚠️ Kein Durchlauf konnte alle APs vollständig zuweisen.")
            else:
                st.error("❌ Keine gültige Lösung gefunden.")
    else:
        st.info("⬅️ Bitte Firmenname eingeben, beide Dateien hochladen und sicherstellen, dass die Firma im AP enthalten ist.")


if __name__ == "__main__":
    main()
