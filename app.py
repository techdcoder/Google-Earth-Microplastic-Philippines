import ee
from generator import MapGenerator
from excel import ExcelInterface
from openpyxl import Workbook
from utilities import *
import folium

class App:
    def __init__(self):
        ee.Authenticate()
        ee.Initialize(project="ee-techdcoder")

        self.philippines = ee.FeatureCollection('projects/ee-techdcoder/assets/philippines-provinces')
        self.province_names = self.philippines.aggregate_array('name').getInfo()
        self.province_count = len(self.province_names)
        self.tags = ['VALUE', 'BINARY', 'DESC']

        self.excel = ExcelInterface('microplastics.xlsx')

    def generate_map(self):
        generator = MapGenerator(self.philippines,[125,22])

        def layer_generator(feature_collection: ee.FeatureCollection, data):

            layers = []
            for column_name in data:
                if column_name == "provinces":
                    continue
                layer_fc = ee.FeatureCollection(feature_collection)

                layer_colors = {}
                if data[column_name]['max'] is None:
                    for i in range(0,len(data[column_name]["data"])):
                        layer_colors[0] = (rgba(120,120,120,120))

                def map_color(feature : ee.Feature):
                    feature = feature.set('style',{
                        "fillColor" :   layer_colors[0]
                    })
                    return feature
                layer_fc = layer_fc.map(map_color).style(styleProperty="style",color="red",width=0.7)
                map_id = layer_fc.getMapId()
                layer = folium.TileLayer(
                    tiles=map_id['tile_fetcher'].url_format,
                    attr="Map Data @copy",
                    overlay=True,
                    name=column_name
                )
                layers.append(layer)
            return layers

        generator.create_layers(layer_generator,self.data)

        generator.generate_html('philippines-microplastics.html')

    def read_excel(self):
        def read_function(workbook: Workbook):
            worksheet = workbook.active
            data = {
                "provinces" : [],
                "microfibers" : {
                    "type": "VALUE",
                    "data": []
                },
                "microbeads" : {
                    "type": "VALUE",
                    "data": []
                },
                "pellets": {
                    "type": "VALUE",
                    "data": []
                },
                "foamed": {
                    "type": "VALUE",
                    "data": []
                },
                "presence": {
                    "type": "BINARY",
                    "data": []
                }
            }
            letter = 'A'
            for column_name in data:
                if column_name == "provinces":
                    continue
                letter = chr(ord(letter) + 1)
                max_value = None
                for row in range(0, self.province_count):
                    row += 2
                    address = letter + str(row)
                    value = worksheet[address].value
                    if value is not None:
                        if max_value is None or value > max_value:
                            max_value = value

                    data[column_name]["data"].append(value)
                    data[column_name]["max"] = max_value
            return data
        self.data = self.excel.read(read_function)
        print(self.data)

    def generate_excel(self):
        def generate_excel_content(workbook: Workbook):
            worksheet = workbook.active

            worksheet.column_dimensions['A'].width = 20
            worksheet.column_dimensions['B'].width = 20
            worksheet.column_dimensions['C'].width = 20
            worksheet.column_dimensions['D'].width = 20
            worksheet.column_dimensions['E'].width = 20
            worksheet.column_dimensions['F'].width = 20
            worksheet.column_dimensions['G'].width = 20

            worksheet['A1'] = 'Province Name'
            row = 2
            for province_name in self.province_names:
                address = 'A' + str(row)
                row += 1
                worksheet[address] = province_name

            # Columns
            worksheet['B1'] = '[VALUE]\nMicrofibers'
            worksheet['C1'] = '[VALUE]\nMicrobeads'
            worksheet['D1'] = '[VALUE]\nPellets'
            worksheet['E1'] = '[VALUE]\nFoamed'
            worksheet['F1'] = '[BINARY]\nPresence'
            worksheet['G1'] = '[DESC]\nComments'

        self.excel.generate(generate_excel_content)
        self.excel.save()
