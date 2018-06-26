"""
Program to automatically calculate all the required data and plot them on a line graph for data analysis
How to use it:
Run the program and enter the directory which contains the CSV Files
Then the program will create another CSV file with all the required data calcualated for each planter and
Write them to excel files and store them in seperate folders based on planter names
There is also the option to view the line graph for a arbitrary time range speified in datetime (YYYY-MM-DD HH:MM:SS)
Created on 06-04-2018 by Sanjeev Rajasekaran
Last Edited on 06-26-2018
"""

import datetime  # Python's datettime module
import os  # import the module which can read file size, get current working directory
import sys  # To exit program if needed
import csv  # import the python module to be able write our data to a CSV file
import xlsxwriter  # Module to write csv data to excel file
import pandas as pd  # Plotting module
import matplotlib.pyplot as plt  # Plotting module
import seaborn as sns  # Plotting Module
import glob  # Get files of specific type
import pathlib  # Module to Get current dir
from collections import defaultdict # Module for pythons defaultdict module

sns.set() # Set Seaborn attributes to default

# Global Constants

DEPTH_BELOW_INVERT_DAN = 0.48
DEPTH_BELOW_INVERT_SAM = 0.50
DEPTH_BELOW_INVERT_SARAH = 0.50
DEPTH_BELOW_INVERT_DYLAN = 0.44
DIAMETER_OUTFLOW_PIPE_DAN = 0.60
DIAMETER_OUTFLOW_PIPE_SAM = 0.60
DIAMETER_OUTFLOW_PIPE_SARAH = 0.60
DIAMETER_OUTFLOW_PIPE_DYLAN = 0.59
SQUARE_ROOT_2GD_PLANTER_DAN = 6.23
SQUARE_ROOT_2GD_PLANTER_SAM = 6.21
SQUARE_ROOT_2GD_PLANTER_SARAH = 6.23
SQUARE_ROOT_2GD_PLANTER_DYLAN = 6.15
WEIRD_EQUATION_CD = 0.53
MONITORING_CHAMBER_SURFACE = 3.14
CONVERSION_TO_CUBIC_FEET = 28.32  # ft^3 - L
DRAINAGE_AREA_PLANTER_DAN = 1530.7
DRAINAGE_AREA_PLANTER_SAM = 1172.1
DRAINAGE_AREA_PLANTER_SARAH = 585.5
DRAINAGE_AREA_PLANTER_DYLAN = 1059.3
PLANTER_SURFACE_AREA = 72


class CalcColumns:

    def __init__(self):

        # Define all the variables

        self.rain_inches_list = list()
        self.depth_above_invert_list = list()
        self.outflow_cfs_list = list()
        self.outflow_ls_list = list()
        self.cumulative_rain_list = list()
        self.final_values_list = list()
        self.cumulative_outflow_volume_list = list()
        self.cumulative_inflow_volume_roof_list = list()
        self.cumulative_inflow_ord_list = list()
        self.event_list = list()

        # Initialise the Count Variables

        self.cumulative_rain_count = 0
        self.outflow_cfs_count = 0
        self.cumulative_outflow_count = 0
        self.roof_method_inflow_cfs_count = 0
        self.cumulative_inflow_volume_roof_method_count = 0
        self.cumulative_inflow_ord_l_count = 0
        self.rain_event_counter_dan = 1
        self.rain_event_counter_sam = 1
        self.rain_event_counter_sarah = 1
        self.rain_event_counter_dylan = 1

        self.columns = ["         Date          ", "Depth outflow (cm)", "Depth outflow (feet)", "Depth above invert",
                        "Depth Diameter",
                        "Outflow CFS", "Outflow LS", "Cumulative outflow volume", "Rain (mm)", "Rain (inch)",
                        "Cumulative Rain", "   Roof Method Inflow   ", "   Roof Method Inflow LS   ",
                        "   Cumulative Inflow Volume Roof Method   ", "Max Depth of Outflow (inch)"]

    """
    All the functions below use formulas from the excel sheet. We add all the values to a list to do future calculations
    """

    def calc_depth_of_outflow_ft(self, depth_of_outflow_cm, planter):

        depth_of_outflow_ft = list()

        for item in depth_of_outflow_cm[planter]:
            depth_of_outflow_ft.append(round((float(item) / 30.48), 2))

        return depth_of_outflow_ft

    def calc_depth_above_invert_ft(self, depth_of_outflow_ft, PLANTER_NAME):

        if PLANTER_NAME == "DAN":
            global DEPTH_BELOW_INVERT_DAN
            DEPTH_BELOW_INVERT = DEPTH_BELOW_INVERT_DAN
        if PLANTER_NAME == "SAM":
            global DEPTH_BELOW_INVERT_SAM
            DEPTH_BELOW_INVERT = DEPTH_BELOW_INVERT_SAM
        if PLANTER_NAME == "SARAH":
            global DEPTH_BELOW_INVERT_SARAH
            DEPTH_BELOW_INVERT = DEPTH_BELOW_INVERT_SARAH
        if PLANTER_NAME == "DYLAN":
            global DEPTH_BELOW_INVERT_DYLAN
            DEPTH_BELOW_INVERT = DEPTH_BELOW_INVERT_DYLAN

        for item in depth_of_outflow_ft:
            self.depth_above_invert_list.append(round((item - DEPTH_BELOW_INVERT), 4))

        return self.depth_above_invert_list

    def calc_depth_diameter(self, depth_above_invert, PLANTER_NAME):

        if PLANTER_NAME == "DAN":
            global DIAMETER_OUTFLOW_PIPE_DAN
            DIAMETER_OUTFLOW_PIPE = DIAMETER_OUTFLOW_PIPE_DAN

        if PLANTER_NAME == "SAM":
            global DIAMETER_OUTFLOW_PIPE_SAM
            DIAMETER_OUTFLOW_PIPE = DIAMETER_OUTFLOW_PIPE_SAM

        if PLANTER_NAME == "SARAH":
            global DIAMETER_OUTFLOW_PIPE_SARAH
            DIAMETER_OUTFLOW_PIPE = DIAMETER_OUTFLOW_PIPE_SARAH

        if PLANTER_NAME == "DYLAN":
            global DIAMETER_OUTFLOW_PIPE_DYLAN
            DIAMETER_OUTFLOW_PIPE = DIAMETER_OUTFLOW_PIPE_DYLAN

        depth_diameter = list()

        for item in depth_above_invert:
            depth_diameter.append(round((item / DIAMETER_OUTFLOW_PIPE), 4))

        return depth_diameter

    def calculate_outflow_cfs(self, depth_above_invert, depth_diameter, depth_of_outflow, date, PLANTER_NAME):

        global WEIRD_EQUATION_CD
        global MONITORING_CHAMBER_SURFACE

        if PLANTER_NAME == "DAN":
            global SQUARE_ROOT_2GD_PLANTER_DAN
            global DIAMETER_OUTFLOW_PIPE_DAN

            SQUARE_ROOT_2GD = SQUARE_ROOT_2GD_PLANTER_DAN
            DIAMETER_OUTFLOW_PIPE = DIAMETER_OUTFLOW_PIPE_DYLAN

        if PLANTER_NAME == "SAM":
            global SQUARE_ROOT_2GD_PLANTER_SAM
            global DIAMETER_OUTFLOW_PIPE_SAM
            SQUARE_ROOT_2GD = SQUARE_ROOT_2GD_PLANTER_SAM
            DIAMETER_OUTFLOW_PIPE = DIAMETER_OUTFLOW_PIPE_SAM

        if PLANTER_NAME == "SARAH":
            global SQUARE_ROOT_2GD_PLANTER_SARAH
            global DIAMETER_OUTFLOW_PIPE_SARAH
            SQUARE_ROOT_2GD = SQUARE_ROOT_2GD_PLANTER_SARAH
            DIAMETER_OUTFLOW_PIPE = DIAMETER_OUTFLOW_PIPE_SARAH

        if PLANTER_NAME == "DYLAN":
            global SQUARE_ROOT_2GD_PLANTER_DYLAN
            SQUARE_ROOT_2GD = SQUARE_ROOT_2GD_PLANTER_DYLAN
            DIAMETER_OUTFLOW_PIPE = DIAMETER_OUTFLOW_PIPE_DYLAN

        for item in depth_diameter:

            if self.outflow_cfs_count == 0:  # If it is the first element, then just add 0, otherwise do rest of logic

                self.outflow_cfs_list.append(0)
                self.outflow_cfs_count += 1
                continue

            if depth_above_invert[self.outflow_cfs_count] > 0:
                outflow_cfs = (2 / 3) * SQUARE_ROOT_2GD * pow(DIAMETER_OUTFLOW_PIPE,
                                                              2) * WEIRD_EQUATION_CD * pow(item, 1.87)
            else:


                try:
                    outflow_cfs = round((MONITORING_CHAMBER_SURFACE * abs(
                    float(depth_of_outflow[self.outflow_cfs_count]) - float(
                        depth_of_outflow[self.outflow_cfs_count - 1])) / c1.calculate_hours_minutes(
                    date[self.outflow_cfs_count], date[self.outflow_cfs_count - 1])), 4)
                except IndexError:
                    print(date)
                    print(self.outflow_cfs_count)

            self.outflow_cfs_list.append(outflow_cfs)
            self.outflow_cfs_count += 1

        return self.outflow_cfs_list

    def calculate_hours_minutes(self, date_current, date_prev):

        date_prev = str(date_prev).strip('"')
        date_current = str(date_current).strip('"')

        try:
            date_prev = datetime.datetime.strptime(date_prev, '%d-%m-%Y %H:%M')  # Use pythons timedelta module
            date_current = datetime.datetime.strptime(date_current,
                                                      '%d-%m-%Y %H:%M')  # to calculate difference between dates
        except ValueError:
            date_prev = datetime.datetime.strptime(date_prev, '%Y-%m-%d %H:%M:%S').strftime('%d-%m-%Y %H:%M')
            date_current = datetime.datetime.strptime(date_current, '%Y-%m-%d %H:%M:%S').strftime('%d-%m-%Y %H:%M')
            date_prev = datetime.datetime.strptime(date_prev, '%d-%m-%Y %H:%M')
            date_current = datetime.datetime.strptime(date_current, '%d-%m-%Y %H:%M')
        time_delta_value = ((date_current - date_prev).total_seconds())

        return time_delta_value

    def calculate_outflow_ls(self, outflow_cfs):

        for item in outflow_cfs:
            self.outflow_ls_list.append(round((item * CONVERSION_TO_CUBIC_FEET), 4))

        return self.outflow_ls_list

    def calculate_cumulative_outflow_volume(self, outflow_cfs, date):

        for item in outflow_cfs:

            if self.cumulative_outflow_count == 0:
                self.cumulative_outflow_volume_list.append(0)
                self.cumulative_outflow_count += 1
                continue
            try:
                self.cumulative_outflow_volume_list.append(self.cumulative_outflow_volume_list[
                                                               self.cumulative_outflow_count - 1] + item * c1.calculate_hours_minutes(
                    date[self.cumulative_outflow_count], date[self.cumulative_outflow_count - 1]))
                self.cumulative_outflow_count += 1
            except IndexError:
                print(date)

        return self.cumulative_outflow_volume_list

    def calculate_rain_inches(self, rain_mm):

        for item in rain_mm:
            self.rain_inches_list.append(round((float(item) / 25.4), 4))

        return self.rain_inches_list

    def calculate_cumulative_rain(self, rain_mm):

        for item in rain_mm:
            self.cumulative_rain_list.append(round((sum(self.rain_inches_list[:(self.cumulative_rain_count + 1)])), 2))
            self.cumulative_rain_count += 1

        return self.cumulative_rain_list

    def calculate_roof_method_inflow_cfs(self, rain_inch, date, PLANTER_NAME):

        if PLANTER_NAME == "DAN":
            global DRAINAGE_AREA_PLANTER_DAN
            DRAINAGE_AREA_PLANTER = DRAINAGE_AREA_PLANTER_DAN

        if PLANTER_NAME == "SAM":
            global DRAINAGE_AREA_PLANTER_SAM
            DRAINAGE_AREA_PLANTER = DRAINAGE_AREA_PLANTER_SAM

        if PLANTER_NAME == "SARAH":
            global DRAINAGE_AREA_PLANTER_SARAH
            DRAINAGE_AREA_PLANTER = DRAINAGE_AREA_PLANTER_SARAH

        if PLANTER_NAME == "DYLAN":
            global DRAINAGE_AREA_PLANTER_DYLAN
            DRAINAGE_AREA_PLANTER = DRAINAGE_AREA_PLANTER_DYLAN

        roof_method_inflow_cfs_list = list()

        for item in rain_inch:
            roof_method_inflow_cfs_list.append(round(
                ((float(item) / 12) * DRAINAGE_AREA_PLANTER) / c1.calculate_hours_minutes(
                    date[self.roof_method_inflow_cfs_count], date[self.roof_method_inflow_cfs_count - 1]), 4))
            self.roof_method_inflow_cfs_count += 1

        return roof_method_inflow_cfs_list

    def calculate_roof_method_inflow_ls(self, inflow_cfs):

        global CONVERSTION_TO_CUBIC_FEET

        roof_method_inflow_ls = list()

        for item in inflow_cfs:
            roof_method_inflow_ls.append(round((item * CONVERSION_TO_CUBIC_FEET), 4))

        return roof_method_inflow_ls

    def calculate_cumulative_inflow_volume_roof_method(self, roof_method_inflow_cfs, date):

        for item in roof_method_inflow_cfs:

            if self.cumulative_inflow_volume_roof_method_count == 0:
                self.cumulative_inflow_volume_roof_list.append(0)
                self.cumulative_inflow_volume_roof_method_count += 1
                continue

            self.cumulative_inflow_volume_roof_list.append(self.cumulative_inflow_volume_roof_list[
                                                               self.cumulative_inflow_volume_roof_method_count - 1] + item * c1.calculate_hours_minutes(
                date[self.cumulative_inflow_volume_roof_method_count],
                date[self.cumulative_inflow_volume_roof_method_count - 1]))
            self.cumulative_inflow_volume_roof_method_count += 1

        return self.cumulative_inflow_volume_roof_list

    def calculate_cumulative_inflow_ORD_L(self, ord_inflow, date):

        for item in ord_inflow:

            if self.cumulative_inflow_ord_l_count == 0:
                self.cumulative_inflow_ord_list.append(0)
                self.cumulative_inflow_ord_l_count += 1
                continue

            # if self.cumulative_inflow_ord_l_count == 2390:
            #   break

            self.cumulative_inflow_ord_list.append(round((self.cumulative_inflow_ord_list[
                                                              self.cumulative_inflow_ord_l_count - 1] + ((float(
                item) + float(ord_inflow[self.cumulative_inflow_ord_l_count - 1])) / 2) * c1.calculate_hours_minutes(
                date[self.cumulative_inflow_ord_l_count], date[self.cumulative_inflow_ord_l_count - 1])), 3))
            self.cumulative_inflow_ord_l_count += 1

        return self.cumulative_inflow_ord_list

    def calculate_max_depth_outflow(self):

        max_depth_outflow_feet = max(self.depth_above_invert_list)
        max_depth_outflow_inch = max_depth_outflow_feet * 12

        return max_depth_outflow_feet, max_depth_outflow_inch

    def calculate_peak_rainfall(self):

        peak_rainfall = max(self.rain_inches_list)

        return peak_rainfall

    def calculate_peak_flow_cfs(self):

        peak_flow_cfs = max(self.outflow_cfs_list)

        return peak_flow_cfs

    def calculate_peak_flow_ls(self):

        peak_flow_ls = max(self.outflow_ls_list)

        return peak_flow_ls

    def calculate_rainflow_inflow_volume(self):

        return max(self.cumulative_inflow_volume_roof_list)  # Return the highest number from the list

    def calculate_outflow_volume(self):

        return max(self.cumulative_outflow_volume_list)

    def calculate_total_rainfall(self):

        return max(self.cumulative_rain_list)

    def calculate_rainfall_volume_x_roof(self, total_rainfall, PLANTER_NAME):

        if PLANTER_NAME == "DAN":
            global DRAINAGE_AREA_PLANTER_DAN
            DRAINAGE_AREA_PLANTER = DRAINAGE_AREA_PLANTER_DAN

        if PLANTER_NAME == "SAM":
            global DRAINAGE_AREA_PLANTER_SAM
            DRAINAGE_AREA_PLANTER = DRAINAGE_AREA_PLANTER_SAM

        if PLANTER_NAME == "SARAH":
            global DRAINAGE_AREA_PLANTER_SARAH
            DRAINAGE_AREA_PLANTER = DRAINAGE_AREA_PLANTER_SARAH

        if PLANTER_NAME == "DYLAN":
            global DRAINAGE_AREA_PLANTER_DYLAN
            DRAINAGE_AREA_PLANTER = DRAINAGE_AREA_PLANTER_DYLAN

        return total_rainfall / 12 * DRAINAGE_AREA_PLANTER

    def calculate_ord_inflow_volume(self, total_rainfall):

        global CONVERSTION_TO_CUBIC_FEET
        global PLANTER_SURFACE_AREA

        return max(self.cumulative_inflow_ord_list) / CONVERSTION_TO_CUBIC_FEET + (
                total_rainfall / 12) * PLANTER_SURFACE_AREA

    def create_csv(self, final_dict):

        """
        Takes 1 Dictionary as argument
        Read the values from the dictionary and pass them on to write to a csv file
        """

        date = list()
        depth_outflow_cm_planter_dan = list()
        depth_outflow_cm_planter_sam = list()
        depth_outflow_cm_planter_sarah = list()
        depth_outflow_cm_planter_dylan = list()
        rain_mm = list()
        final_dict_dates = sorted(final_dict)  # Sort the dates

        for key in final_dict_dates:
            date.append(key)
            depth_outflow_cm_planter_dan.append(final_dict[key]['p1'])
            depth_outflow_cm_planter_sam.append(final_dict[key]['p2'])
            depth_outflow_cm_planter_sarah.append(final_dict[key]['p3'])
            depth_outflow_cm_planter_dylan.append(final_dict[key]['p4'])
            rain_mm.append(final_dict[key]['rain'])

        c1.detect_rain_events(final_dict)  # pass the dictionary to detect rain events
        c1.write_to_csv(date, depth_outflow_cm_planter_dan, depth_outflow_cm_planter_sam,   # write values to 'data.csv'
                        depth_outflow_cm_planter_sarah, depth_outflow_cm_planter_dylan, rain_mm)

    def detect_rain_events(self, data, gap=12):

        """
        Takes a dictionary and a optional parameter as arguments
        Detects the occurences of rain based on the gap value
        Default value of Gap is 12
        """
        cre_date = list()  # cre stands for Current Rain Event
        cre_depth_outflow = defaultdict(float)
        cre_rain = list()  # cre stands for Current Rain Event
        current_dir = pathlib.Path(__file__).parent
        os.chdir(current_dir)
        dates = sorted(data.keys())
        min_delta = datetime.timedelta(hours=gap)  # min gap with no rain for rain event
        state = 'find_start'
        start_dt = None  # beginning of this rain event
        last_dt = None  # last date rain > 0

        cre_depth_outflow["DAN"] = []
        cre_depth_outflow["SAM"] = []
        cre_depth_outflow["SARAH"] = []
        cre_depth_outflow["DYLAN"] = []
        event_counter_rain = 1
        event_counter_dry = 1

        for dt in dates:

            if state == 'find_end':
                if dt - last_dt >= min_delta: # if the time between the last date of rain and current date is more than the gap value

                    if data[dt]['rain'] == 0:
                        # found the end of the event
                        for planter in ["DAN", "SAM", "SARAH", "DYLAN"]:
                            c1.calc_data(cre_date, cre_depth_outflow, cre_rain, planter)
                        state = 'find_start'
                        start_dt = None
                        last_dt = None
                        cre_date.clear()        # clear all lists
                        cre_rain.clear()
                        cre_depth_outflow["DAN"] = []
                        cre_depth_outflow["SAM"] = []
                        cre_depth_outflow["SARAH"] = []
                        cre_depth_outflow["DYLAN"] = []
                        event_counter_dry += 1  # increment the event counters
                        event_counter_rain += 1
                    else:
                        last_dt = dt
                        cre_date.append(str(dt))
                        cre_rain.append(data[dt]['rain'])
                        event = "Rain" + str(event_counter_dry)
                        self.event_list.append(event)
                        cre_depth_outflow["DAN"].append(data[dt]['p1'])
                        cre_depth_outflow["SAM"].append(data[dt]['p2'])
                        cre_depth_outflow["SARAH"].append(data[dt]['p3'])
                        cre_depth_outflow["DYLAN"].append(data[dt]['p4'])

                elif data[dt]['rain'] > 0:
                    # found another rain value in current rain event
                    last_dt = dt
                    cre_date.append(str(dt))
                    cre_rain.append(data[dt]['rain'])
                    event = "Rain" + str(event_counter_rain)
                    self.event_list.append(event)
                    cre_depth_outflow["DAN"].append(data[dt]['p1'])
                    cre_depth_outflow["SAM"].append(data[dt]['p2'])
                    cre_depth_outflow["SARAH"].append(data[dt]['p3'])
                    cre_depth_outflow["DYLAN"].append(data[dt]['p4'])

                else:
                    cre_date.append(str(dt))
                    cre_rain.append(data[dt]['rain'])
                    event = "Rain" + str(event_counter_dry)
                    self.event_list.append(event)
                    cre_depth_outflow["DAN"].append(data[dt]['p1'])
                    cre_depth_outflow["SAM"].append(data[dt]['p2'])
                    cre_depth_outflow["SARAH"].append(data[dt]['p3'])
                    cre_depth_outflow["DYLAN"].append(data[dt]['p4'])

            if state == 'find_start' and data[dt]['rain'] > 0:
                # found the beginning of a new rain event
                state = 'find_end'  # next, look for the end of the rain event
                start_dt = dt  # start of the current rain event
                last_dt = dt  # most recent time that we've seen rain
                cre_date.append(str(dt))
                cre_rain.append(data[dt]['rain'])
                cre_depth_outflow["DAN"].append(data[dt]['p1'])
                cre_depth_outflow["SAM"].append(data[dt]['p2'])
                cre_depth_outflow["SARAH"].append(data[dt]['p3'])
                cre_depth_outflow["DYLAN"].append(data[dt]['p4'])
                event = "Rain" + str(event_counter_rain)
                self.event_list.append(event)
            if state == 'find_start' and data[dt]['rain'] == 0:
                event = "Dry" + str(event_counter_dry)
                self.event_list.append(event)

        # ran out of data
        if state == 'find-end':
            # currently looking for the end of a rain event so yield what we have
            for planter in ["DAN", "SAM", "SARAH", "DYLAN"]:
                c1.calc_data(cre_date, cre_depth_outflow, cre_rain, planter)

    def calc_data(self, date, depth_outflow_cm, rain_mm, planter):
        """
        Call all funtions with the required arguments and when its done, call the function to write the
        calculated data to a excel file
        """

        depth_outflow_feet = c1.calc_depth_of_outflow_ft(depth_outflow_cm, planter)
        depth_above_invert = c1.calc_depth_above_invert_ft(depth_outflow_feet, planter)
        depth_diamter = c1.calc_depth_diameter(depth_above_invert, planter)
        outflow_cfs = c1.calculate_outflow_cfs(depth_above_invert, depth_diamter, depth_outflow_feet, date,
                                               planter)
        outflow_ls = c1.calculate_outflow_ls(outflow_cfs)
        cumulative_outflow_volume = c1.calculate_cumulative_outflow_volume(outflow_cfs, date)
        rain_inches = c1.calculate_rain_inches(rain_mm)
        cumulative_rain = c1.calculate_cumulative_rain(rain_mm)
        roof_method_inflow_cfs = c1.calculate_roof_method_inflow_cfs(rain_inches, date, planter)
        roof_method_inflow_ls = c1.calculate_roof_method_inflow_ls(roof_method_inflow_cfs)
        cumulative_inflow_volume_roof_method = c1.calculate_cumulative_inflow_volume_roof_method(roof_method_inflow_cfs,
                                                                                                 date)
        # cumulative_inflow = c1.calculate_cumulative_inflow_ORD_L(ord_inflow_ls, date)
        max_depth_outflow_feet, max_depth_outflow_inch = c1.calculate_max_depth_outflow()
        peak_rainfall = c1.calculate_peak_rainfall()
        peak_flow_cfs = c1.calculate_peak_flow_cfs()
        peak_flow_ls = c1.calculate_peak_flow_ls()
        rainfall_inflow_volume = c1.calculate_rainflow_inflow_volume()
        outflow_volume = c1.calculate_outflow_volume()
        total_rainfall = c1.calculate_total_rainfall()
        rainfall_volume_x_roof = c1.calculate_rainfall_volume_x_roof(total_rainfall, planter)
        # ord_inflow_volume = c1.calculate_ord_inflow_volume(total_rainfall)
        c1.write_to_excel(date, depth_outflow_cm, depth_outflow_feet, depth_above_invert, depth_diamter, outflow_cfs,
                          outflow_ls, cumulative_outflow_volume, rain_mm, rain_inches, cumulative_rain,
                          roof_method_inflow_cfs, roof_method_inflow_ls, cumulative_inflow_volume_roof_method,
                          max_depth_outflow_feet, max_depth_outflow_inch, peak_rainfall, peak_flow_cfs, peak_flow_ls,
                          rainfall_inflow_volume, outflow_volume, total_rainfall, rainfall_volume_x_roof, planter)
        c1.clear_all()

    def clear_all(self):

        """
        Clears all lists and counter variables
        """
        self.rain_inches_list.clear()
        self.depth_above_invert_list.clear()
        self.outflow_cfs_list.clear()
        self.outflow_ls_list.clear()
        self.cumulative_rain_list.clear()
        self.final_values_list.clear()
        self.cumulative_outflow_volume_list.clear()
        self.cumulative_inflow_volume_roof_list.clear()
        self.cumulative_inflow_ord_list.clear()
        self.cumulative_rain_count = 0
        self.outflow_cfs_count = 0
        self.cumulative_outflow_count = 0
        self.roof_method_inflow_cfs_count = 0
        self.cumulative_inflow_volume_roof_method_count = 0
        self.cumulative_inflow_ord_l_count = 0

    def write_to_csv(self, date, depth_outflow_cm_planter_dan, depth_outflow_cm_planter_sam,
                     depth_outflow_cm_planter_sarah, depth_outflow_cm_planter_dylan, rain_mm):

        """
        Write the values into data.csv file
        """

        event_list = self.event_list
        current_dir = pathlib.Path(__file__).parent
        os.chdir(current_dir)

        myfile = open('Data.csv', 'w', newline='')
        with myfile:
            writer = csv.writer(myfile)
            writer.writerow(["Date", "Rain (mm)", "Depth Outflow P0", "Depth Outflow P1", "Depth Outflow P2",
                             "Depth Outflow P3", "Event"])
            for date_cur, rain_mm_cur, depth_outflow_cm_cur_dan, depth_outflow_cm_cur_sam, \
                depth_outflow_cm_cur_sarah, depth_outflow_cm_cur_dylan, event in zip(date, rain_mm,
                                                                              depth_outflow_cm_planter_dan,
                                                                              depth_outflow_cm_planter_sam,
                                                                              depth_outflow_cm_planter_sarah,
                                                                              depth_outflow_cm_planter_dylan, event_list):

                writer.writerow([date_cur, rain_mm_cur, depth_outflow_cm_cur_dan,
                                 depth_outflow_cm_cur_sam, depth_outflow_cm_cur_sarah, depth_outflow_cm_cur_dylan,
                                 event])

    def write_to_excel(self, date, depth_outflow_cm, depth_outflow_feet, depth_above_invert, depth_diamter, outflow_cfs,
                       outflow_ls, cumulative_outflow_volume, rain_mm, rain_inches, cumulative_rain,
                       roof_method_inflow_cfs, roof_method_inflow_ls, cumulative_inflow_volume_roof_method,
                       max_depth_outflow_feet, max_depth_outflow_inch, peak_rainfall, peak_flow_cfs, peak_flow_ls,
                       rainfall_inflow_volume, outflow_volume, total_rainfall, rainfall_volume_x_roof, PLANTER_NAME):

        """
         Write the contents of our calculations to a Excel file row by row
        """
        csv_row_count = 0
        final_values_counter = 0
        row = 0
        col = 0

        self.final_values_list = ["Max Depth of Outflow (feet)", max_depth_outflow_feet,
                                  "Max Depth of Outflow (inch)",
                                  max_depth_outflow_inch, "Peak Rainfall", peak_rainfall, "Peak Flow CFS",
                                  peak_flow_cfs, "Peak Flow LS", peak_flow_ls, "Rainfall Inflow Volume",
                                  rainfall_inflow_volume, "Outflow Volume", outflow_volume, "Total Rainfall",
                                  total_rainfall, "Rainfall volume * Roof", rainfall_volume_x_roof]

        current_dir = pathlib.Path(__file__).parent

        if PLANTER_NAME == "DAN":
            directory = str(current_dir) + "/" + PLANTER_NAME  # check if directory exists, if not
            if not os.path.exists(directory):                  # create it and change current directory to
                os.makedirs(directory)                         # the created directory
            os.chdir(directory)
            current_file_name = "RainGraph" + str(self.rain_event_counter_dan) + "-" + PLANTER_NAME + ".xlsx"
            workbook = xlsxwriter.Workbook(current_file_name)
            self.rain_event_counter_dan += 1

        if PLANTER_NAME == "SAM":
            directory = str(current_dir) + "/" + PLANTER_NAME
            if not os.path.exists(directory):
                os.makedirs(directory)
            os.chdir(directory)
            current_file_name = "RainGraph" + str(self.rain_event_counter_sam) + "-" + PLANTER_NAME + ".xlsx"
            workbook = xlsxwriter.Workbook(current_file_name)
            self.rain_event_counter_sam += 1

        if PLANTER_NAME == "SARAH":
            directory = str(current_dir) + "/" + PLANTER_NAME
            if not os.path.exists(directory):
                os.makedirs(directory)
            os.chdir(directory)
            current_file_name = "RainGraph" + str(self.rain_event_counter_sarah) + "-" + PLANTER_NAME + ".xlsx"
            workbook = xlsxwriter.Workbook(current_file_name)
            self.rain_event_counter_sarah += 1

        if PLANTER_NAME == "DYLAN":
            directory = str(current_dir) + "/" + PLANTER_NAME
            if not os.path.exists(directory):
                os.makedirs(directory)
            os.chdir(directory)
            current_file_name = "RainGraph" + str(self.rain_event_counter_dylan) + "-" + PLANTER_NAME + ".xlsx"
            workbook = xlsxwriter.Workbook(current_file_name)
            self.rain_event_counter_dylan += 1

        worksheet = workbook.add_worksheet("Data")

        cell_format_center = workbook.add_format({'align': 'center'})
        # cell_format.set_align('center')

        c1.set_column_width(worksheet, self.columns)

        worksheet.write(row, col, "Date", cell_format_center)
        worksheet.write(row, col + 1, "Depth outflow (cm)", cell_format_center)
        worksheet.write(row, col + 2, "Depth outflow (feet)", cell_format_center)
        worksheet.write(row, col + 3, "Depth above invert", cell_format_center)
        worksheet.write(row, col + 4, "Depth Diameter", cell_format_center)
        worksheet.write(row, col + 5, "Outflow CFS", cell_format_center)
        worksheet.write(row, col + 6, "Outflow LS", cell_format_center)
        worksheet.write(row, col + 7, "Cumulative outflow volume", cell_format_center)
        worksheet.write(row, col + 8, "Rain (mm)", cell_format_center)
        worksheet.write(row, col + 9, "Rain (inch)")
        worksheet.write(row, col + 10, "Cumulative Rain", cell_format_center)
        # worksheet.write(row, col + 11, "ORD Inflow LS", cell_format_center)
        worksheet.write(row, col + 11, "Roof Method Inflow", cell_format_center)
        worksheet.write(row, col + 12, "Roof Method Inflow LS", cell_format_center)
        worksheet.write(row, col + 13, "Cumulative Inflow Volume Roof Method", cell_format_center)

        row += 1

        for item_date in date:

            if final_values_counter <= 17:

                worksheet.write(row, col, item_date, cell_format_center)
                worksheet.write(row, col + 1, depth_outflow_cm[csv_row_count], cell_format_center)
                worksheet.write(row, col + 2, depth_outflow_feet[csv_row_count], cell_format_center)
                worksheet.write(row, col + 3, depth_above_invert[csv_row_count], cell_format_center)
                worksheet.write(row, col + 4, depth_diamter[csv_row_count], cell_format_center)
                worksheet.write(row, col + 5, outflow_cfs[csv_row_count], cell_format_center)
                worksheet.write(row, col + 6, outflow_ls[csv_row_count], cell_format_center)
                worksheet.write(row, col + 7, cumulative_outflow_volume[csv_row_count], cell_format_center)
                worksheet.write(row, col + 8, rain_mm[csv_row_count], cell_format_center)
                worksheet.write(row, col + 9, rain_inches[csv_row_count], cell_format_center)
                worksheet.write(row, col + 10, cumulative_rain[csv_row_count], cell_format_center)
                # worksheet.write(row, col + 11, ord_inflow_ls[csv_row_count], cell_format_center)
                worksheet.write(row, col + 11, roof_method_inflow_cfs[csv_row_count], cell_format_center)
                worksheet.write(row, col + 12, roof_method_inflow_ls[csv_row_count], cell_format_center)
                worksheet.write(row, col + 13, cumulative_inflow_volume_roof_method[csv_row_count], cell_format_center)
                # worksheet.write(row, col + 14, cumulative_inflow[csv_row_count], cell_format_center)
                worksheet.write(row, col + 14, self.final_values_list[final_values_counter], cell_format_center)
                worksheet.write(row, col + 15, self.final_values_list[final_values_counter + 1], cell_format_center)

                row += 1
                csv_row_count += 1
                final_values_counter += 2

            else:

                worksheet.write(row, col, item_date, cell_format_center)
                worksheet.write(row, col + 1, depth_outflow_cm[csv_row_count], cell_format_center)
                worksheet.write(row, col + 2, depth_outflow_feet[csv_row_count], cell_format_center)
                worksheet.write(row, col + 3, depth_above_invert[csv_row_count], cell_format_center)
                worksheet.write(row, col + 4, depth_diamter[csv_row_count], cell_format_center)
                worksheet.write(row, col + 5, outflow_cfs[csv_row_count], cell_format_center)
                worksheet.write(row, col + 6, outflow_ls[csv_row_count], cell_format_center)
                worksheet.write(row, col + 7, cumulative_outflow_volume[csv_row_count], cell_format_center)
                worksheet.write(row, col + 8, rain_mm[csv_row_count], cell_format_center)
                worksheet.write(row, col + 9, rain_inches[csv_row_count], cell_format_center)
                worksheet.write(row, col + 10, cumulative_rain[csv_row_count], cell_format_center)
                # worksheet.write(row, col + 11, ord_inflow_ls[csv_row_count], cell_format_center)
                worksheet.write(row, col + 11, roof_method_inflow_cfs[csv_row_count], cell_format_center)
                worksheet.write(row, col + 12, roof_method_inflow_ls[csv_row_count], cell_format_center)
                worksheet.write(row, col + 13, cumulative_inflow_volume_roof_method[csv_row_count], cell_format_center)
                # worksheet.write(row, col + 14, cumulative_inflow[csv_row_count], cell_format_center)

                row += 1
                csv_row_count += 1

        if len(date) <= 17:

            final_values_remaining = len(self.final_values_list) - row * 2
            start_value = len(self.final_values_list) - (final_values_remaining // 2 + 1)
            for final_list_value in self.final_values_list[start_value::]:

                worksheet.write(row, col + 14, self.final_values_list[final_values_counter], cell_format_center)
                try:
                    worksheet.write(row, col + 15, self.final_values_list[final_values_counter + 1], cell_format_center)
                except IndexError:
                    break
                row += 1
                final_values_counter += 2

        ws = workbook.add_worksheet("Graph")

        # ---OPTION TO ADD CHART AS IMAGE VIA SCRIPT RATHER THAN GENERATING IT IN EXCEL---
        # img = openpyxl.drawing.image.Image('april_storm_dan.png')
        # ws.add_image(img)
        # ---END---

        date_formula = '=Data!$A$2:$A$' + str(row)
        value_formula_outflow = '=Data!$G$2:$G$' + str(row)
        value_formula_roof = '=Data!$M$2:$M$' + str(row)
        chart1 = workbook.add_chart({'type': 'line'})
        chart1.add_series({'categories': date_formula, 'values': value_formula_outflow, 'line': {'width': 4}})
        chart1.add_series({'categories': date_formula, 'values': value_formula_roof, 'line': {'width': 4}})
        chart1.set_title({'name': 'Outflow Vs Roof Method Inflow'})
        chart1.set_x_axis({'name': 'Time', 'text_axis': True, 'date_axis': False, 'label_position': 'low'})
        chart1.set_y_axis({'name': 'Flow L/S'})
        # Set an Excel chart style. Colors with white outline and shadow.
        chart1.set_style(10)

        # Insert the chart into the worksheet (with an offset).
        ws.insert_chart('F3', chart1)
        chart1.set_size({'width': 750, 'height': 400})
        workbook.close()

    def line_chart_outflow(self, start_time, end_time):

        current_dir = pathlib.Path(__file__).parent
        os.chdir(current_dir)
        line_chart_date_list = list()
        line_chart_rain_mm_list = list()
        line_chart_depth_outflow_list_dan = list()
        line_chart_depth_outflow_list_sam = list()
        line_chart_depth_outflow_list_sarah = list()
        line_chart_depth_outflow_list_dylan = list()
        plot_data_running = 0

        try:
            with open("Data.csv", "r+") as fp:

                lines = fp.readlines()
                for line in lines[1:]:
                    line = line.strip()
                    split_words = line.split(',')
                    date = split_words[0]
                    rain_mm = split_words[1]
                    depth_outflow_dan = split_words[2]
                    depth_outflow_sam = split_words[3]
                    depth_outflow_sarah = split_words[4]
                    depth_outflow_dylan = split_words[5]

                    if date == start_time:
                        plot_data_running = 1
                    if plot_data_running == 1:
                        line_chart_date_list.append(date)
                        line_chart_rain_mm_list.append(rain_mm)
                        line_chart_depth_outflow_list_dan.append(depth_outflow_dan)
                        line_chart_depth_outflow_list_sam.append(depth_outflow_sam)
                        line_chart_depth_outflow_list_sarah.append(depth_outflow_sarah)
                        line_chart_depth_outflow_list_dylan.append(depth_outflow_dylan)
                    if date == end_time:
                        plot_data_running = 0
                        break

        except FileNotFoundError:
            print("Cannot open file")
            sys.exit()

        workbook = xlsxwriter.Workbook("Outflow_datum.xlsx")
        worksheet = workbook.add_worksheet("Data")
        row = 0
        col = 0
        # cell_format.set_align('center')
        cell_format_center = workbook.add_format({'align': 'center'})
        worksheet.write(row, col, "Date", cell_format_center)
        worksheet.write(row, col + 1, "Rain (mm)", cell_format_center)
        worksheet.write(row, col + 2, "Dan Outflow", cell_format_center)
        worksheet.write(row, col + 3, "Sam Outflow", cell_format_center)
        worksheet.write(row, col + 4, "Sarah Outflow", cell_format_center)
        worksheet.write(row, col + 5, "Dylan Outflow", cell_format_center)
        worksheet.write(row, col + 6, "Cumulative Inflow", cell_format_center)
        row += 1
        cumulative_inflow = c1.cumulative(line_chart_rain_mm_list)
        for date_iter, rain_iter, outflow_dan, outflow_sam, outflow_sarah, outflow_dylan, inflow in zip(line_chart_date_list,
                                                                                                line_chart_rain_mm_list,
                                                                                                line_chart_depth_outflow_list_dan,
                                                                                                line_chart_depth_outflow_list_sam,
                                                                                                line_chart_depth_outflow_list_sarah,
                                                                                                line_chart_depth_outflow_list_dylan, cumulative_inflow):
            worksheet.write(row, col, date_iter, cell_format_center)
            worksheet.write(row, col+1, float(rain_iter), cell_format_center)
            worksheet.write(row, col + 2, float(outflow_dan), cell_format_center)
            worksheet.write(row, col + 3, float(outflow_sam), cell_format_center)
            worksheet.write(row, col + 4, float(outflow_sarah), cell_format_center)
            worksheet.write(row, col + 5, float(outflow_dylan), cell_format_center)
            worksheet.write(row, col + 6, float(inflow), cell_format_center)
            row += 1

        value_formula_rain = '=Data!$B$2:$B$' + str(row)
        value_formula_dan = '=Data!$C$2:$C$' + str(row)
        value_formula_sam = '=Data!$D$2:$D$' + str(row)
        value_formula_sarah = '=Data!$E$2:$E$' + str(row)
        value_formula_dylan = '=Data!$F$2:$F$' + str(row)
        value_formula_cumulative_rain = '=Data!$G$2:$G$' + str(row)
        min_list = list()
        max_list = list()
        for list_item in [line_chart_depth_outflow_list_dan, line_chart_depth_outflow_list_sam, line_chart_depth_outflow_list_sarah, line_chart_depth_outflow_list_dylan]:
            list_item = list(map(float, list_item))
            min_list.append(min(list_item))
            max_list.append(max(list_item))

        min_value = min(min_list)
        max_value = max(max_list)

        for sheet in ["Planter Outflow", "Cumulative Inflow"]:
            ws = workbook.add_worksheet(sheet)
            # cell_format.set_align('center')
            chart1 = workbook.add_chart({'type': 'line'})
            date_formula = '=Data!$A$2:$A$' + str(row)
            if sheet == "Planter Outflow":
                chart1.add_series({'categories': date_formula, 'values': value_formula_dan, 'line': {'width': 4}})
                chart1.add_series({'categories': date_formula, 'values': value_formula_sam, 'line': {'width': 4}})
                chart1.add_series({'categories': date_formula, 'values': value_formula_sarah, 'line': {'width': 4}})
                chart1.add_series({'categories': date_formula, 'values': value_formula_dylan, 'line': {'width': 4}})
                chart1.set_title({'name': 'Outflow Datum '})
                chart1.set_x_axis({'name': 'Time', 'text_axis': True, 'date_axis': False, 'label_position': 'low'})
                chart1.set_y_axis({'name': 'Inflow vs Outflow', 'min': min_value, 'max': max_value})
                # Set an Excel chart style. Colors with white outline and shadow.
                chart1.set_style(10)
                # Insert the chart into the worksheet (with an offset).
                ws.insert_chart('F3', chart1)
                chart1.set_size({'width': 750, 'height': 400})

            if sheet == "Cumulative Inflow":

                chart1.add_series({'categories': date_formula, 'values': value_formula_cumulative_rain, 'line': {'width': 4}})
                chart1.set_title({'name': 'Cumulative Inflow'})
                chart1.set_x_axis({'name': 'Time', 'text_axis': True, 'date_axis': False, 'label_position': 'low'})
                chart1.set_y_axis({'name': 'Inflow'})
                # Set an Excel chart style. Colors with white outline and shadow.
                chart1.set_style(10)
                # Insert the chart into the worksheet (with an offset).
                ws.insert_chart('F3', chart1)
                chart1.set_size({'width': 750, 'height': 400})
        workbook.close()

    def cumulative(self, seq):
        """ Given a sequence, return a list of values of the cumulative """
        total = 0
        result = list()
        for item in seq:
            total += float(item)
            result.append(total)
        return result

    def set_column_width(self, worksheet, columns):
        length_list = [len(x) for x in columns]
        for i, width in enumerate(length_list):
            worksheet.set_column(i, i, width)


c1 = CalcColumns()  # Object of our class


def main():

    data = defaultdict(lambda: defaultdict(float))

    input_path = input("Enter Directory of CSV Files:")  # Get the path from the user

    if len(os.listdir(input_path)) == 0:
        print("Directory is empty")
        sys.exit()

    else:

        for file in glob.glob(os.path.join(input_path, 'StevensA[0-9]*.csv')):

            try:
                fp = open(file, 'r')  # Try to open it in "read-only" mode
            except FileNotFoundError:  # If we run into any problems, Throw an exception and return the below message
                print("File {} cannot be opened".format(file))
                sys.exit(0)
            with fp:

                if os.stat(file).st_size != 0:  # check if file is empty
                    for line, row in enumerate(fp):
                        if line in [0, 1, 2, 3]:
                            continue  # skip the header row
                        cols = row.strip().split(',')
                        dt = datetime.datetime.strptime(cols[0].strip('"'), "%Y-%m-%d %H:%M:%S")
                        data[dt]['rain'] = float(cols[172])
                        data[dt]['p1'] = float(cols[6])
                        data[dt]['p2'] = float(cols[12])
                        data[dt]['p3'] = float(cols[18])
                        data[dt]['p4'] = float(cols[24])

                else:
                    continue
        c1.create_csv(data)

    line_chart = input("Do you want to see outflow line chart for a time period? y/n \n")

    if line_chart.lower() == "y":
        time_start = input("Enter start time period \n")
        time_end = input("Enter End time period \n")
        c1.line_chart_outflow(time_start, time_end)


if __name__ == '__main__':
    main()
