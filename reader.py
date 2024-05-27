from openpyxl import load_workbook

class Position:
    def __init__(self, column='A', row=1):
        self.row = row
        self.column = column

    def __repr__(self):
        return self.to_str()

    def increase_column(self, n):
        self.column = chr(ord(self.column) + n)

    def decrease_column(self, n):
        self.column = chr(ord(self.column) - n)

    def increase_row(self, n):
        self.row += n

    def decrease_row(self, n):
        self.row -= n
    def to_str(self):
        return self.column + str(self.row)

    def copy(self):
        return Position(self.column,self.row)

def is_tag(cell_value: str):
    cell_value = str(cell_value)
    if cell_value is None:
        return False
    if len(cell_value) < 3:
        return False
    return (cell_value[0] == '[') and (cell_value[-1] == ']')


def remove_end_spaces(text):
    while text[0] == ' ':
        text = text[1:]
    while text[-1] == ' ':
        text = text[:-1]
    return text
def get_tag(tag_str: str):
    tag_str = tag_str[1:][:-1]
    if ":" in tag_str:
        tag,data = tag_str.split(":")
        tag,data = remove_end_spaces(tag),remove_end_spaces(data)
        return tag,data
    return tag_str, None

def get_table_header(worksheet, position):
    position = position.copy()
    table_header = {
        "ELEMENTS" : 0,
        "PROPERTIES" : 0,
        "SRC" : None,
        "NAME" : None,
        "MAPTYPE" : None
    }
    position.increase_row(1)
    while True:
        cell_content = worksheet[position.to_str()].value
        if not is_tag(cell_content):
            break
        tag, data = get_tag(cell_content)
        if tag == "END":
            break
        if tag in table_header:
            table_header[tag] = data
        else:
            raise Exception(f"Invalid Tag: {tag}")
        position.increase_column(1)
    return table_header

def verify_property_list(properties):
    types = [
        "END",
        "ID",
        "VALUE",
        "DESC",
        "BOOL",
        "LONGITUDE",
        "LATITUDE",
        "CLASSIFICATION"
    ]

    if len(properties) == 0:
        raise Exception("No properties!")

    id_exists = False

    for name in properties:
        property_data = properties[name]
        if property_data is None:
            raise Exception("Invalid Tag!")
        if property_data['TYPE'] == 'ID':
            id_exists = True
    if not id_exists:
        raise Exception("No ID Property!")


def get_id_column(properties):
    id_column = None
    for name in properties:
        property = properties[name]
        if property is None:
            continue

        type = property["TYPE"]
        if type == "ID":
            return property["COLUMN"]
    raise Exception("No ID column in properties!")

def get_element_count(worksheet, properties):
    id_column = get_id_column(properties)
    position = Position(id_column, 4)

    element_count = 0
    while worksheet[position.to_str()].value is not None:
        position.increase_row(1)
        element_count += 1
    return element_count

def get_elements(worksheet,header, properties, table_position: Position):
    map_type = header["MAPTYPE"]

    element_count = get_element_count(worksheet,properties)
    position = table_position.copy()
    position.increase_row(3)

    elements = {}
    for _ in range(0, element_count):
        element = {}
        id = None

        for name in properties:
            property = properties[name]

            column = property['COLUMN']
            type = property['TYPE']

            property_position = Position(column, position.row)
            cell_value = worksheet[property_position.to_str()].value

            if type == "ID":
                id = cell_value

            element[name] = cell_value

        if map_type == "MARKER-MAP":
            if not  id in elements:
                elements[id] = []
            elements[id].append(element)
        else:
            elements[id] = element
        position.increase_row(1)
    return elements
def get_properties(worksheet, position : Position):
    position = position.copy()
    properties = {}
    position.increase_row(2)

    while True:
        cell_str = worksheet[position.to_str()].value

        if cell_str is None:
            break

        if is_tag(cell_str):
            type, data = get_tag(cell_str)
            column = position.column

            if type == "END":
                break

            if data is None:
                properties[type] = {
                    "TYPE": type,
                    "COLUMN": column
                }
            else:
                properties[data] = {
                    "TYPE": type,
                    "COLUMN" : column
                }
            position.increase_column(1)
        else:
            raise Exception("Property is not formatted as a tag!")

    if len(properties) == 0:
        raise Exception("No properties!")

    verify_property_list(properties)
    return properties

def verify_header(header):
    if header["NAME"] is None:
        raise Exception("Table name is not defined")
    if header["MAPTYPE"] is None:
        raise Exception("Map Type is not defined")

def read_file(file_path):
    workbook = load_workbook(file_path)
    worksheet = workbook.active

    scan_range = 100
    tables = {}
    for column in range(ord('A'),ord('Z')+1):
        column = str(chr(column))
        for row in range(1, scan_range):
            position = Position(column,row)

            address = position.to_str()
            cell_content = worksheet[address].value
            if is_tag(cell_content):
                tag,data = get_tag(cell_content)

                if tag == "TABLE":
                    table_header = get_table_header(worksheet,position)
                    verify_header(table_header)

                    properties = get_properties(worksheet,position)

                    elements = get_elements(worksheet, table_header, properties, position)

                    table_name = table_header['NAME']
                    tables[table_name] = {
                        "HEADER": table_header,
                        "DATA": {
                            "PROPERTIES": properties,
                            "ELEMENTS": elements
                        }
                    }
    return tables

def get_properties_from_table(table):
    return table['DATA']['PROPERTIES']

def get_value_from_table(table, id_name, property_name):
    elements = table['DATA']['ELEMENTS']
    return elements[id_name][property_name]

def get_values_from_table(table, property_name):
    elements = table['DATA']['ELEMENTS']
    values = {}
    for id in elements:
        values[id] = get_value_from_table(table, id, property_name)
    return values

def get_source_from_table(table):
    return table['HEADER']['SRC']