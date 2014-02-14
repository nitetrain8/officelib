"""
Created on Jan 13, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather
"""

from pstats import Stats

stats_file = "C:\\Users\\PBS Biotech\\Documents\\Personal\\PBS_Office\\MSOffice\\officelib\\pbslib\\test\\profile2.txt"

from datetime import datetime

s = Stats(stats_file)
s.strip_dirs()
s.sort_stats('time')

s.print_callers(0.1)


