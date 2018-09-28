# firebird-selector
#
#REQUIREMENTS:
# pip install fdb pandas xlwt argparse pywin32

# -*- coding: utf-8 -*-


#########################   CONFIG   ###########################
# Config for Firebird DB connector
# Carefully fill in this section!
db = 'path_to_db, WindowsExample: 192.192.192.1:c:\base\base.gdb'
dbuser = 'user'
dbpass = 'password'
#########################   CONFIG   ###########################

import sys
import fdb
import pandas as pd
import argparse

# description of the script
desc = """
fb-selector allows us to select PASS CARD id and numbers 
of employees of a specified department or organization
"""

# chosen query for target requests
query = """
select distinct
        p.tableno as tn,
	p.name ||' '|| p.firstname ||' '|| p.secondname as fio,        
        c.sitecode||c.cardno as card_code
        
from
        person p
JOIN dictvals d on
        %s
JOIN pass on
        pass.personid = p.personid
JOIN card c on
        c.cardid = pass.cardid
where
        d.attributeval like "%%%s%%"
        and c.cardstatus = 1
"""



	
def connectdb():
	return fdb.connect(
		dsn = db,
		user = dbuser, 
		password = dbpass,
		sql_dialect = 1, # urgently necessary, without it is unable to process the query
		charset = 'WIN1251'
		)

		
def prepquery(a,q):
	if a.dep is None:
		join_condition = 'd.dictvalid = p.orgid'
		name = a.org
	else:
		join_condition = 'd.dictvalid = p.depid'
		name= a.dep
	
	return q % (join_condition, name)

	
def query_and_save(q, o_file):
	# connect to database
	con = connectdb()
	
	# read SQL query right into dataframe 
	df = pd.read_sql_query(q, con)
	out = pd.ExcelWriter(o_file)
	df.to_excel(out)
	out.save()
	
	#close
	con.close()
	out.close()

def autofit(o_file):
	#Excel cell width autofit
	try:
		import win32com.client as win32
		excel = win32.gencache.EnsureDispatch('Excel.Application')
		wb = excel.Workbooks.Open(r'%s' % o_file)
		ws = wb.Worksheets("Sheet1")
		ws.Columns.AutoFit()
		wb.Save()
		excel.Application.Quit()
		print ("Cells width in output file is autofitted successfully ")
	except Exception as e:
		print ('ERROR! Width of cells in outfile was not autofitted!\n ErrorMessage: \n%s' % str(e))


#:::::::::::::::::::::::::   SCRIPT   :::::::::::::::::::::::::#

####################   Reading Arguments   #####################		
parser = argparse.ArgumentParser(description=desc)
group = parser.add_mutually_exclusive_group()
group.add_argument('-d', '--depatrment', dest='dep', help='Submit department name')
group.add_argument('-o', '--organization', dest='org',	help='Submit organization name')
parser.add_argument('-out', '--output', dest='out', default='out.xls',
					help='Submit output Excel file name. (default is out.xls)')
args = parser.parse_args()
####################   Reading Arguments   #####################	

#######################   Script BODY   ########################
	
if args.dep is None and args.org is None:
	# exit script - none of department or organization were submitted
	sys.exit("None of department or organization were submitted - nothing to do. Read help -h")	
else:
	q = prepquery(args, query)
	print ("QUERY: \n", q)
	query_and_save(q, args.out)
	autofit(args.out)
	

#######################   Script BODY   ########################
