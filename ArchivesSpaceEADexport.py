from asnake.client import ASnakeClient
import asnake.client
import json
import re
import html
import re
import csv
import xlsxwriter
import getopt
import sys
import os
from getpass import getpass
import colorist
from colorama import Fore, Back, Style

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

def makeExport(ubl_num, client):
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
	currentitems = 0

	search_eads = client.get('repositories/2/search', params={'q': 'identifier:' + ubl_num, 'type': ['resource'], 'page': 1}).json()
	if search_eads['total_hits'] == 0:
		print(Fore.RED + "ERROR:" + Style.RESET_ALL + " Geen resultaten gevonden voor " + ubl_num + ".")
		#input("Druk op een knop om dit venster te sluiten")
		return()

	this_ead = json.loads(search_eads['results'][0]['json'])
	print(ubl_num + " gevonden op locatie " + this_ead['uri'])

	#json.dumps zodat python json doorgeeft ipv python code
	#Search eerst een keer om te zien hoeveel resultaten er zijn, en hoeveel verdere queries er gemaakt moeten worden
	search_results = client.get('repositories/2/search', params={'aq': json.dumps(makeJsonQuery(this_ead['uri'])) , 'type': ['archival_object'], 'page': 1, 'page_size': 250}).json()
	#print("First page: " + str(search_results['first_page']) + " - Last page: " + str(search_results['last_page']) + " - This page: " + str(search_results['this_page']) + " - First item in set: " + str(search_results['offset_first']) + " - Last item in set: " + str(search_results['offset_last']))

	if search_results['total_hits'] > 0:
		for pagenmbr in range(search_results['last_page']):
			#Doe meer queries voor volgende pagina's indien die er zijn.
			if search_results['this_page'] != search_results['last_page']:
				search_results = client.get('repositories/2/search', params={'aq': json.dumps(makeJsonQuery(this_ead['uri'])) , 'type': ['archival_object'], 'page': pagenmbr +1, 'page_size': 250}).json()
				#print("First page: " + str(search_results['first_page']) + " - Last page: " + str(search_results['last_page']) + " - This page: " + str(search_results['this_page']) + " - First item in set: " + str(search_results['offset_first']) + " - Last item in set: " + str(search_results['offset_last']))
			print("Bezig met verwerken pagina " + str(search_results['this_page']) + " van " + str(search_results['last_page']) + ", item " + str(search_results['offset_first']) + " tot en met " + str(search_results['offset_last']))
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
							if 'type' in my_record['notes'][a-1]:
								if my_record['notes'][a-1]['type'] == 'physdesc':
									for b in range((len(my_record['notes'][a-1]['content']))):
										physdescvar = physdescvar + my_record['notes'][a-1]['content'][b-1] + "^"
						worksheet.write(row, 5, physdescvar[:-1])
						worksheet.write(row, 7, odd_var[:-1])
						worksheet.write(row, 8, subnotesvar[:-1])
						row+=1

	workbook.close()

	print(Fore.GREEN + "Export succesvol afgerond!" + Style.RESET_ALL + " Uw export is te vinden op " + os.getcwd() + '\\output_' + ubl_num + '.xlsx')

wwfile = os.getcwd() + "\\aspw.txt"

print("--Welkom bij de ArchivesSpace EAD export module--")
print("--Deze module werkt alleen als u gebruik maakt van een PC die verbonden is aan het UB netwerk--")
print("--Dit betreft vaste PC's aangesloten op het UB met een kabel, of apparaten verbonden met NUWD-laptop--")
print("--Indien u EduVPN gebruikt zal deze module niet werken. Zet dus eerste EduVPN uit!--")

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
try:
	client = ASnakeClient(baseurl="https://collectionguides.universiteitleiden.nl/staff/api",
						  username=authun,
						  password=authww)
	client.authorize()
except asnake.client.web_client.ASnakeAuthError as e:
	print(Fore.RED + "ERROR:" + Style.RESET_ALL + " Kan geen verbinding maken met ArchivesSpace API")
	print(Fore.RED + "ERROR:" + Style.RESET_ALL + " Zorg er voor dat u een computer gebruikt die aan het UBL netwerk verbonden is, en controlleer uw inloggegevens")
	input("Druk op een knop om dit venster te sluiten")
	raise SystemExit(e)

ubl_numb = ""

while ubl_numb != "exit":
	ubl_numb = "ubl" + input("Voer UBL nummer in (zonder \"ubl\"), of voer \"exit\" in om af te sluiten: ")
	if ubl_numb == "ublexit":
		sys.exit(2)
	else:
		makeExport(ubl_numb, client)

