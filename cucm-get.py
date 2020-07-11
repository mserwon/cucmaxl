import os
from ciscoaxl import axl
import json
import zeep.helpers
from pprintjson import pprintjson
#from py_dotenv import read_dotenv

#dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
#read_dotenv(dotenv_path)

cucm = '10.22.192.11'
cucm_username = 'admin'
cucm_password = 'Pr3s!d!0'
cucm_version = '12.0'

ucm = axl(username=cucm_username,password=cucm_password,cucm=cucm,cucm_version=cucm_version)

#cucm_response = ucm.get_user('hwilliams')
cucm_response = ucm.get_users()
#JSON Fixup for response
#cucm_response.replace("'", '"')
#cucm_response.replace("None","'None'")
cucm_in_dict = zeep.helpers.serialize_object(cucm_response)
cucm_json = json.loads(json.dumps(cucm_in_dict))

cucmfile = open('cucm-get.json','w')
pprintjson(cucm_json,file=cucmfile)

cucmfile.close()