from simple_salesforce import Salesforce
from dataclasses import dataclass, asdict
from openpyxl import Workbook
from dotenv import dotenv_values
import json
import csv
import requests
requests.packages.urllib3.disable_warnings()
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


@dataclass
class FieldDefinition:

    is_custom: bool
    default_value: str
    dependent_picklist: str
    digits: int
    is_encrypted: bool
    is_externalid: bool
    help_text: str
    label: str
    length: int
    name: str
    is_namefield: bool
    picklist_values: list[str]
    precision: int
    lookup_to: str
    field_type: str

@dataclass
class ObjectDefinition:
    name: str
    label: str
    metadata_type: str
    is_custom: str
    fields: list[FieldDefinition]


ACCESS_TOKEN = ''
INSTANCE_URL = ''

def run():

    config = dotenv_values('.env')

    # Execute
    global ACCESS_TOKEN
    global INSTANCE_URL

    ACCESS_TOKEN, INSTANCE_URL = authenticate(config)

    sobject_data = get_sobjects()

    object_definitions = parse_sobjects(sobject_data)

    # object_definitions_with_details = add_object_details(object_definitions[0:5])
    object_definitions_with_details = add_object_details(object_definitions)

    all_object_data = {}

    for object_definition in object_definitions_with_details:
        
        data_as_rows = build_data_rows(object_definition)
        all_object_data[object_definition.name] = data_as_rows
    

    build_xlsx(all_object_data)
    # build_csv(data_as_rows)


def create_field_definition(data: str):

    picklist_data = ''
    lookup_to = ''

    # Build picklelist string
    if data['type'] == 'picklist':
        for picklist_value in data['picklistValues']:
            if picklist_value['active']:
                picklist_as_string = f"{picklist_value['label']} ({picklist_value['value']})"
                picklist_data += picklist_as_string + '\n'

    # Lookup data
    if data['type'] == 'reference':
        if data['referenceTo']:
            lookup_to = data['referenceTo'][0]



    field_obj = FieldDefinition(data['custom'],
                                data['defaultValue'],
                                data['dependentPicklist'],
                                data['digits'],
                                data['encrypted'],
                                data['externalId'],
                                data['inlineHelpText'],
                                data['label'],
                                data['length'],
                                data['name'],
                                data['nameField'],
                                picklist_data,
                                data['precision'],
                                lookup_to,
                                data['type']
                                )
    
    return field_obj


def create_object_definition(data: str, metadata_type: str):

    obj = ObjectDefinition(
        data['name'],
        data['label'],
        metadata_type,
        data['custom'],
        None
    )

    return obj


def authenticate(config):

        sf = Salesforce(username=config['USERNAME'], password=config['PASSWORD'],
                 consumer_key= config['CLIENT_KEY'], consumer_secret= config['CLIENT_SECRET'], domain='test')

        
        # return access_token, instance_url
        return sf.session_id, sf.base_url

def get_sobjects():
    
    GET_OBJECTS_METHOD = 'sobjects/'

    url = INSTANCE_URL+GET_OBJECTS_METHOD
    headers = {'Authorization' : 'Bearer ' + ACCESS_TOKEN}
    
    r = requests.get(url=url, headers=headers, verify=False)

    return r.json()


def parse_sobjects(data: str):

    # Rules

    # | TYPE              |   Creatable   | Triggerable   | Deletable | Custom    |
    # | ------------------|---------------|---------------|-----------|-----------|
    # | Standard Object   |   true        | true          | true      | false     |
    # | Custom Object     |   true        | true          | true      | true      |
    # | Custom Setting    |   true        | false         | true      | true      |
    # | Custom Metadata   |   false       | false         | false     | true      |
    # | Platform Event    |   true        | true          | false     | true      |


    TRIGGERABLE = 'triggerable'
    CREATABLE = 'createable'
    DELETABLE = 'deletable'
    CUSTOM = 'custom'

    object_definitions = []
    
    for sobject in data['sobjects']:

        object_name = sobject['name']
        object_type = ''

        if '__mdt' in object_name:
            object_type = 'Custom Metadata'

        elif '__e' in object_name:
            object_type = 'Platform Event'

        elif sobject[CREATABLE] == True and sobject[DELETABLE] == True:
            if sobject[TRIGGERABLE] == True:
                if sobject[CUSTOM] == False:
                    object_type = 'Standard Object'
                else:
                    object_type = 'Custom Object'

            elif sobject[TRIGGERABLE] == False and sobject[CUSTOM] == True:
                object_type = 'Custom Setting'

        if object_type != '':
            object_definitions.append(create_object_definition(sobject, object_type))
    
    return object_definitions


def add_object_details(object_definitions :list[ObjectDefinition]):

    for object_definition in object_definitions:

        object_describe_data = get_sobject_describe(object_definition.name)

        all_field_data = []

        for field_data in object_describe_data['fields']:
            all_field_data.append(create_field_definition(field_data))
        
        object_definition.fields = all_field_data
    
    return object_definitions



def get_sobject_describe(object_name: str):

    GET_DESCRIBE = f'sobjects/{object_name}/describe'

    headers = {'Authorization' : 'Bearer ' + ACCESS_TOKEN}
    url_describe = INSTANCE_URL+GET_DESCRIBE
    r_describe = requests.get(url=url_describe, headers=headers, verify=False)

    return r_describe.json()


def get_object_map():

    object_map = {}
    object_map['label'] = 'Object: Label'
    object_map['name'] = 'Object: Name'
    object_map['metadata_type'] = 'Object: Type'
    object_map['is_custom'] = 'Object: IsCustom'

    return object_map

def get_field_map():

    field_map = {}
    field_map['label'] = 'Field: Label'
    field_map['name'] = 'Field: Name'
    field_map['is_custom'] = 'Field: IsCustom'
    field_map['help_text'] = 'Field: HelpText'
    field_map['field_type'] = 'Field: Type'
    field_map['length'] = 'Field: Length'
    field_map['lookup_to'] = 'Field: Lookup To'
    field_map['picklist_values'] = 'Field: Picklist Values'
    field_map['default_value'] = 'Field: Default Value'
    field_map['dependent_picklist'] = 'Field: Dependent Picklist'
    field_map['digits'] = 'Field: Digits'
    field_map['precision'] = 'Field: Precision'
    field_map['is_encrypted'] = 'Field: IsEncrypted'
    field_map['is_externalid'] = 'Field: IsExternalId'
    field_map['is_namefield'] = 'Field: IsNameField'

    return field_map


def get_header_array():

    object_column_headers = list(get_object_map().values())
    field_column_headers = list(get_field_map().values())

    return object_column_headers + field_column_headers

    
def convert_object_definition_to_array(object_definition: ObjectDefinition):

    object_definition_dict = asdict(object_definition)

    object_keys = list(get_object_map().keys())
    field_keys = list(get_field_map().keys())

    # Object Values
    base_row = []
    for key in object_keys:
        base_row.append(str(object_definition_dict[key]))

    all_rows = []

    # Field Values
    for field_definition in object_definition_dict['fields']:
        
        row = []

        for key in field_keys:
            row.append(str(field_definition[key]))
        
        all_rows.append(base_row+row)
    
    return all_rows


def build_data_rows(object_definition: ObjectDefinition):

    rows = convert_object_definition_to_array(object_definition)

    return rows


def build_csv(data_as_rows):

    with open('object_data_file.csv', mode='w') as object_data_file:

        csv_writer = csv.writer(object_data_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)

        csv_writer.writerow(get_header_array())

        for row in data_as_rows:
            csv_writer.writerow(row)


def build_xlsx(all_objects_data):


    wb = Workbook()

    for object_name in all_objects_data.keys():

        ws = wb.create_sheet(object_name)
        ws.append(get_header_array())

        for row in all_objects_data[object_name]:
            for cell in row:
                cell = ILLEGAL_CHARACTERS_RE.sub(r'', cell)
            ws.append(row)

    wb.save('output.xlsx')


run()



