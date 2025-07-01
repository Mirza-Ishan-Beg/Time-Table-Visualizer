## Table of Content

1. [[#Installation]]
2. [[#Usage]]

---

## Installation

This works on windows (and on Linux with some tweaks, make sure you know how to install without messing your system pips up.), with `.xlsx` files to fetch the time table.

Get the code locally by downloading and unzipping the files.

Install the files in `requirements.txt` using the command `pip install requirements.txt`

You should be now good to go.

---

## Usage

- The `.xlsx` should be named `Weekly_Schedule.xlsx` (you can change the name within `TimeTableVisualizer.py` if you want, but this name is hard coded so you should know what you are doing while editing and making further changes.) 
- The excel sheet needs to have `hhmm - hhmm` time stamps within a specific day section. 
- There also needs to be `timetables_chart` directory on the exact hierarchical level of `TimeTableVisualizer.py` and the `Weekly_Schedule.xlsx`
- The excel sheet here is an example of how time, days, subjects and their codes should be structured in order for the script to work.
- The visuals it presents are in a 24 hours clock.

---
