import os
from ciscoaxl import axl
import json
import zeep.helpers
import pandas as pd
from pprintjson import pprintjson
#from py_dotenv import read_dotenv

#dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
#read_dotenv(dotenv_path)

cucm = '10.22.192.11'
cucm_username = 'admin'
cucm_password = 'Pr3s!d!0'
cucm_version = '12.0'

ucm = axl(username=cucm_username,password=cucm_password,cucm=cucm,cucm_version=cucm_version)

cucm_response = ucm.sql_query('select * from enduser')['row']
#JSON Fixup for response
#cucm_response.replace("'", '"')
#cucm_response.replace("None","'None'")
cucm_in_dict = zeep.helpers.serialize_object(cucm_response)
col = []
for x in cucm_in_dict[0]:
    col.append(x.tag)
df = pd.DataFrame(columns=col)

row = []
for obj in cucm_in_dict[1:]:
    row = {}
    for attrib in obj:
        row[attrib.tag] = attrib.text
    df = df.append(row, True)
#cucm_json = json.loads(json.dumps(cucm_in_dict))

cucmfile = open('cucm-sql.csv','w')
# pprintjson(cucm_json,file=cucmfile)
df.to_csv(cucmfile, index= False, header=True)

cucmfile.close()