from dateutil.relativedelta import relativedelta
import datetime
from typing import List
from worker import WorkPackage, Worker
from typing import Optional, Dict
import math
round_num = 4

def safe_floor_division(max_assignable, step):
    res = 0
    while res <= max_assignable:
        res += step
    return res

def step_per_month(h_per_month, months, required_pm, step):
    max_h = round(safe_floor_division(h_per_month, step), round_num)
    ap_per_m = []
    for month in months:
        if required_pm >= sum(ap_per_m) + max_h:
            ap_per_m.append(max_h)
        elif required_pm < sum(ap_per_m) + max_h:
            ap_per_m.append(round(required_pm - sum(ap_per_m), round_num))
        else:
            ap_per_m.append(0)
    return ap_per_m

def assign_ap_to_workers(work_package: WorkPackage, workers: List[Worker], start_year: int):
    if math.isnan(work_package.required_pm):
        return None

    if work_package.assigned_worker_id is not None:
        result = assign_to_fixed_worker(work_package, workers, start_year)
        if result is not None:
            return result  # fixed assignment complete

    required_pm = work_package.required_pm
    assignment = {}
    start_dt = datetime.datetime.strptime(work_package.start_date, "%d.%m.%Y")
    end_dt = datetime.datetime.strptime(work_package.end_date, "%d.%m.%Y")

    work_package.assigned_worker_name = ""
    work_package.assigned_worker_ids = []
    remaining_pm = required_pm
    worker_contributions = {}

    # Lista de todos os meses do AP
    months = []
    current_dt = start_dt
    while current_dt <= end_dt:
        months.append(current_dt)
        current_dt += relativedelta(months=1)

    h_per_month = work_package.required_pm / len(months)

    for w in sorted(workers, key=lambda w: -w.salary):
        if remaining_pm <= 0:
            break

        step = w.step
        medium_step_per_month = step_per_month(h_per_month, months, required_pm, step)

        sum_per_year = 0
        working_non_stop = True
        assigned = False
        sum_pm_this_round_stil = 0

        possible_start = 0
        while working_non_stop:
            current_chunk = []
            current_pm = 0.0
            for i in range(possible_start, len(months)):
                month_dt = months[i]
                year_idx = month_dt.year - start_year
                month_idx = month_dt.month - 1
                if month_dt.month == 0:
                    sum_per_year=0

                if not (0 <= year_idx < len(w.hours_available_per_month)):
                    break

                monthly_available = round(w.hours_available_per_month[year_idx][month_idx], 2)
                yearly_remaining = w.hours_available[year_idx][0] - w.summed_hours_worked_per_year[year_idx] - sum_per_year
                monthly_limit = w.pm_per_ap
                max_assignable = min(medium_step_per_month[i], monthly_available, yearly_remaining, monthly_limit, remaining_pm-current_pm)

                if max_assignable>step:
                    assignable = safe_floor_division(max_assignable, step)
                    assignable = assignable - (assignable-max_assignable)
                else:
                    assignable = max_assignable

                if monthly_available - assignable <0 or yearly_remaining - assignable<0 or monthly_limit- assignable<0 or yearly_remaining- assignable<0:
                    break

                if assignable >= step:
                    current_chunk.append((month_dt, year_idx, month_idx, assignable))
                    current_pm += assignable
                    sum_per_year += assignable
                    assigned = True
                else:
                    if assignable>0.02:
                        current_chunk.append((month_dt, year_idx, month_idx, assignable))
                        current_pm += assignable
                        sum_per_year += assignable
                        assigned = True
                    break

                if current_pm + 1e-6 >= remaining_pm:
                    break

            if assigned:
                for (month_dt, year_idx, month_idx, assignable) in current_chunk:
                    key = (w.id, year_idx, month_idx)
                    assignment.setdefault(key, 0.0)
                    assignment[key] += round(assignable,round_num)

                    w.hours_available_per_month[year_idx][month_idx] -= round(assignable,round_num)
                    w.summed_hours_worked_per_year[year_idx] += round(assignable,round_num)
                    w.summed_hours_worked_total += round(assignable,round_num)

                    remaining_pm -= assignable
                    worker_contributions[w.id] = worker_contributions.get(w.id, 0.0) + round(assignable,round_num)
                    w.assignment_log.setdefault((year_idx, month_idx), [])
                    w.assignment_log[(year_idx, month_idx)].append((work_package.id, round(assignable,round_num)))

                    # Then, during assignment:
                    w.assignment_log_hours_per_ap[work_package.id] = (
                        w.assignment_log_hours_per_ap.get(work_package.id, 0.0) + round(assignable, round_num)
                    )

                required_pm -= current_pm

                if w.id not in work_package.assigned_worker_ids:
                    work_package.assigned_worker_ids.append(w.id)
                    if work_package.assigned_worker_name:
                        work_package.assigned_worker_name += f", {w.name}"
                    else:
                        work_package.assigned_worker_name = w.name

            working_non_stop = False
            assigned = False

            if round(remaining_pm,3) <= 0:
                break

    if round(sum(worker_contributions.values()), 6) < required_pm:
        not_assigned = round(required_pm - sum(worker_contributions.values()), 6)
        key = (0, -1, -1)
        assignment.setdefault(key, 0.0)
        assignment[key] += not_assigned

        if "Nicht zugewiesen" not in work_package.assigned_worker_name:
            if work_package.assigned_worker_name:
                work_package.assigned_worker_name += ", Nicht zugewiesen"
            else:
                work_package.assigned_worker_name = "Nicht zugewiesen"
        if 0 not in work_package.assigned_worker_ids:
            work_package.assigned_worker_ids.append(0)

    return assignment

def assign_to_fixed_worker(
    work_package: WorkPackage,
    workers: List[Worker],
    start_year: int,
    round_num: int = 2
) -> Optional[Dict[tuple, float]]:
    """
    Assigns the work_package to a fixed worker if `assigned_worker_id` is set.
    Updates the worker's availability and logs, and returns the assignment dict.
    If no fixed worker is assigned or found, returns None.
    """

    if work_package.assigned_worker_id is None:
        return None

    # Find worker with matching ID
    worker = next((w for w in workers if w.id == work_package.assigned_worker_id), None)
    if worker is None:
        print(f"⚠️ Worker ID {work_package.assigned_worker_id} not found.")
        return None

    try:
        start_dt = datetime.datetime.strptime(work_package.start_date, "%d.%m.%Y")
        end_dt = datetime.datetime.strptime(work_package.end_date, "%d.%m.%Y")
    except Exception as e:
        print(f"❌ Invalid date in WorkPackage {work_package.id}: {e}")
        return None

    # Build list of months
    months = []
    current_dt = start_dt
    while current_dt <= end_dt:
        months.append(current_dt)
        current_dt += relativedelta(months=1)

    required_pm = work_package.required_pm
    pm_per_month = round(required_pm / len(months), round_num)
    assignment = {}

    for dt in months:
        year_idx = dt.year - start_year
        month_idx = dt.month - 1

        key = (worker.id, year_idx, month_idx)
        assignment[key] = assignment.get(key, 0.0) + pm_per_month

        # Update worker's availability
        worker.hours_available_per_month[year_idx][month_idx] -= pm_per_month
        worker.summed_hours_worked_per_year[year_idx] += pm_per_month
        worker.summed_hours_worked_total += pm_per_month

        # Update logs
        worker.assignment_log.setdefault((year_idx, month_idx), []).append((work_package.id, pm_per_month))
        worker.assignment_log_hours_per_ap[work_package.id] = (
            worker.assignment_log_hours_per_ap.get(work_package.id, 0.0) + pm_per_month
        )

    # Update work package assignment list
    work_package.assigned_worker_ids = [worker.id]
    if not work_package.assigned_worker_name:
        work_package.assigned_worker_name = f"{worker.name} {worker.surname}"

    return assignment
