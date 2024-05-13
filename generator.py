import folium
import ee
from utilities import *

class MapGenerator:
    def __init__(self, feature_collection: ee.FeatureCollection, location: []):
        self.feature_collection = feature_collection

        self.map = folium.Map()
        self.map.location = location


    def create_layers(self, layer_generator, data):
        layers = layer_generator(self.feature_collection, data)
        for layer in layers:
            layer.add_to(self.map)
        self.map.add_child(folium.LayerControl())

    def generate_html(self, filename: str):
        self.map.save(filename)