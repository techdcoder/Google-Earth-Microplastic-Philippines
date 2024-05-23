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
    if cell_value is None:
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
        type = properties[name]
        if type not in types and name != "LATITUDE" and name != "LONGITUDE":
            print(name)
            raise Exception(f"Invalid tag: {type}!")
        if type == "ID":
            if id_exists:
                raise Exception("Multiple ID's are not allowed!")
            id_exists = True

    if not id_exists:
        raise Exception("No ID Property!")
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

            if type == "END":
                break

            if data is None:
                properties[type] = None
            else:
                properties[data] = type
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

                    table_name = table_header['NAME']
                    tables[table_name] = {
                        "HEADER": table_header,
                        "DATA": {
                            "PROPERTIES": properties,
                            "ELEMENTS": [
                            ]
                        }
                    }
    return tables
def main():
    tables = read_file('sample.xlsx')
    print(tables)

if __name__ == "__main__":
    main()