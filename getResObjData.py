from asnake.client import ASnakeClient
import json
import re
import html
import re
import csv
import xlsxwriter
import getopt
import sys
import os

#Backward search voor velden die aan elkaar gekoppeld zijn. Momenteel zoekt naar Digital Objects die aan Archival Object zijn gekoppeld.
def makeJsonQuery(fieldValue):
	x = {
		"query": 
		{
			"jsonmodel_type": "boolean_query",
			"op": "AND",
			"subqueries": 
			[
				{
					"jsonmodel_type": "field_query",
					"negated": False,
					"field": "keyword",
					"value": fieldValue,
					"comparator": "contains"
				}
			]
		},
		"jsonmodel_type": "advanced_query"
	}
	return x

wwfile = os.getcwd() + "\\aspw.txt"

if os.path.exists(wwfile):
	with open(wwfile, "r") as rfile:
		for i, line in enumerate(rfile):
			if i == 0:
				authun = re.sub("\n", "", line)
			if i == 1:
				authww = re.sub("\n", "", line)
else:
	authun = input("Username: ")
	authww = getpass()

#Connect to ArchivesSpace API
client = ASnakeClient(baseurl="https://collectionguides.universiteitleiden.nl/staff/api",
                      username=authun,
                      password=authww)
client.authorize()

#If given --ubl=XXX argument process that EAD. Else process ubl002.
try:
	opts, args = getopt.getopt(sys.argv[1:], 'u', ['ubl='])
	for opt, arg in opts:
		if opt in '--ubl':
			if not arg:
				print("Error: No argument given.")
				sys.exit(2)
			else:
				ubl_num = "ubl" + arg
		else:
			print("Unregocnized argument: " + opt)
	if not opts:
		ubl_num = "ubl002"
except getopt.GetoptError as err:
	print(str(err))
	sys.exit(2)

workbook = xlsxwriter.Workbook(os.getcwd() + '\\output_' + ubl_num + '.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True, 'bottom': True})
row = 0
column = 0

content = ["uri", "ref_id", "level", "title", "unitid", 
                    "physdesc", "unitdate", "general", "subnotes"]

for item in content :
	worksheet.write(row, column, item, bold)
	column+=1
row += 1
column = 0

search_eads = client.get('repositories/2/search', params={'q': 'identifier:' + ubl_num, 'type': ['resource'], 'page': 1}).json()
if search_eads['total_hits'] > 0:
	this_ead = json.loads(search_eads['results'][0]['json'])

print(this_ead['uri'])

#json.dumps zodat python json doorgeeft ipv python code
#Search eerst een keer om te zien hoeveel resultaten er zijn, en hoeveel verdere queries er gemaakt moeten worden
search_results = client.get('repositories/2/search', params={'aq': json.dumps(makeJsonQuery(this_ead['uri'])) , 'type': ['archival_object'], 'page': 1, 'page_size': 250}).json()
print("First page: " + str(search_results['first_page']) + " - Last page: " + str(search_results['last_page']) + " - This page: " + str(search_results['this_page']) + " - First item in set: " + str(search_results['offset_first']) + " - Last item in set: " + str(search_results['offset_last']))

if search_results['total_hits'] > 0:
	for pagenmbr in range(search_results['last_page']):
		#Doe meer queries voor volgende pagina's indien die er zijn.
		if search_results['this_page'] != search_results['last_page']:
			search_results = client.get('repositories/2/search', params={'aq': json.dumps(makeJsonQuery(this_ead['uri'])) , 'type': ['archival_object'], 'page': pagenmbr +1, 'page_size': 250}).json()
		#Doe je ding
		for index in range(search_results['offset_last'] - search_results['offset_first']):
			date_var = ''
			odd_var = ''
			subnotesvar = ''
			extentsvar = ''
			physdescvar = ''
			my_record = json.loads(search_results['results'][index]['json'])
			if my_record['jsonmodel_type'] != 'resource':
				if (my_record['resource']['ref']) == this_ead['uri']:
					if 'uri' in my_record:
						worksheet.write(row, 0, my_record['uri'])
					if 'ref_id' in my_record:
						worksheet.write(row, 1, my_record['ref_id'])
					if 'level' in my_record:
						worksheet.write(row, 2, my_record['level'])
					if 'title' in my_record:
						worksheet.write(row, 3, my_record['title'])
					if 'component_id' in my_record:
						worksheet.write(row, 4, my_record['component_id'])
					for a in range((len(my_record['dates']))):					
						#Is een dict in een list, dus altijd eerst [0] anders wil hij een int!!
						date_var = date_var + my_record['dates'][a-1]['expression'] + "^"
					worksheet.write(row, 6, date_var[:-1])
					for a in range((len(my_record['notes']))):
						if 'subnotes' in my_record['notes'][a-1]:
							if my_record['notes'][a-1]['type'] == 'odd':
								for b in range((len(my_record['notes'][a-1]['subnotes']))):
									odd_var = odd_var + my_record['notes'][a-1]['subnotes'][b-1]['content'] + "^"
							if my_record['notes'][a-1]['type'] == 'scopecontent':
								for b in range((len(my_record['notes'][a-1]['subnotes']))):
									subnotesvar = subnotesvar + my_record['notes'][a-1]['subnotes'][b-1]['content'] + "^"
						if my_record['notes'][a-1]['type'] == 'physdesc':
							for b in range((len(my_record['notes'][a-1]['content']))):
								physdescvar = physdescvar + my_record['notes'][a-1]['content'][b-1] + "^"
					worksheet.write(row, 5, physdescvar[:-1])
					worksheet.write(row, 7, odd_var[:-1])
					worksheet.write(row, 8, subnotesvar[:-1])
					row+=1

workbook.close()
