'''
short script for extracting data from json files
'''


import json
from rich import print

osiris_data:dict = json.load(open('osiris_data.json'))
person_data:list = json.load(open('person_data.json'))
enriched_data = json.load(open('enriched_data.json'))

person_dict = {a.get('input_name'):a for a in person_data}
osiris_data_w_contacts = {}

for code, entry in osiris_data.items():
    contactdetails = {}
    if entry.get('contacts'):
        for contact in entry.get('contacts'):
            details = person_dict.get(contact)
            if details:
                contactdetails[contact] = {
                    'name': details.get('main_name'),
                    'first_name':details.get('other_names')[0],
                    'email': details.get('email'),
                    'faculty': details.get('faculty'),
                    'orgs': details.get('orgs'),
                    'programmes': details.get('programmes'),
                    'people_page': details.get('people_page_url'),
                }
    entry['contacts'] = contactdetails
    osiris_data_w_contacts[code] = entry


print(len(osiris_data_w_contacts))
# store as json

json.dump(osiris_data_w_contacts, open('osiris_data_w_contacts.json', 'w'), indent=4)