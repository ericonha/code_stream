import numpy as np

list_of_workers = []
list_of_av_worker = []


class Worker:
    """
    The Worker class represents an individual worker and manages their availability and salary.
    """

    def __init__(self, id, salary, years, perc_year, name, surname):
        """
        Initialize a Worker object with an ID, salary, and availability for each year.

        Args:
            id (int): The worker's ID.
            salary (float): The worker's salary.
            years (int): The number of years to track availability.
        """
        self.id = id
        self.hours_available = np.zeros((years, 1))
        self.salary = salary
        self.months = np.zeros((years, 1))
        self.hours_available_per_month = np.zeros((years, 12))
        self.perc_year = perc_year
        self.name = name
        self.surname = surname

    def __lt__(self, other):
        """
        Compare workers based on their salaries.

        Args:
            other (Worker): Another Worker object.

        Returns:
            bool: True if self's salary is greater than other's, False otherwise.
        """
        return self.salary > other.salary

    def discount_hours(self, discounted_hours, year):
        """
        Discount worked hours from the available hours for a specific year.

        Args:
            discounted_hours (float): The number of hours worked to discount.
            year (int): The year to discount hours from.
        """
        self.hours_available[year] = self.hours_available[year] - discounted_hours

    def is_available(self, year, month):
        # Check if the worker is available for the given year and month
        if self.hours_available_per_month[year, month] > 0:
            return True
        return False

    def allowed_hours(self, hours_available, months):
        """
        Set the available hours for each year.

        Args:
            hours_available (float): The number of available hours for each year.
            :param hours_available:
            :param months:
        """

        while len(months) > len(hours_available):
            hours_available.append([0])

        while len(hours_available) > len(months):
            months.append(0)

        for index in range(self.months.shape[0]):
            self.months[index] = months[index] / 12

        for index in range(self.hours_available.shape[0]):
            self.hours_available[index][0] = hours_available[index] * self.months[index]

        counter = 0
        year = 0
        for month_in_year in months:
            if counter == 0 & month_in_year < 12:
                self.hours_available_per_month[year][12 - month_in_year:12] = 1
                counter += 1
            else:
                self.hours_available_per_month[year][0:month_in_year] = 1
            year += 1


def add_to_list(workers):
    """
    Add a list of workers to the global list of workers.

    Args:
        workers (list): A list of Worker objects.
    """
    for w in workers:
        list_of_workers.append(w)


def sorte_workers():
    """
    Sort the global list of workers based on their salaries.
    """
    list_of_workers.sort()
