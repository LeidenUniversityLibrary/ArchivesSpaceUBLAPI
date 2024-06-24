from asnake.client import ASnakeClient
import json
import re
import html
import re
import getopt
import sys
import os

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

#If given --obj=XXX argument process that Archival Object. Else process ubl002.
try:
	opts, args = getopt.getopt(sys.argv[1:], 'o', ['obj='])
	for opt, arg in opts:
		if opt in '--obj':
			if not arg:
				print("Error: No argument given.")
				sys.exit(2)
			else:
				obj_num = arg
		else:
			print("Unregocnized argument: " + opt)
	if not opts:
		obj_num = "66508"
except getopt.GetoptError as err:
	print(str(err))
	sys.exit(2)

digital_object = client.get("/repositories/2/archival_objects/" + obj_num)
print(json.dumps(digital_object.json(), indent=4))

