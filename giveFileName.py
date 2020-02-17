#!/usr/bin/env python
# coding: utf-8
# By Grand Chang. in 2020 CNY. Use for download BT_Billing data .xls (but HTML format)
# Try shorten file name input.

import datetime

a_Fname = 'BT_Billing_'
b_Fname = str( datetime.datetime.today().year)
c_Fname = input('Key file date "mmddhhmm"')
d_Fname = '.xls'
filename = a_Fname + b_Fname + c_Fname + d_Fname
print (filename)