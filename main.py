import ee
import folium

from reader import read_file, get_properties_from_table, get_values_from_table, get_source_from_table
from utilities import rgba, red_hue
if __name__ == "__main__":
    ee.Authenticate()
    ee.Initialize(project="ee-techdcoder")

    tables = read_file("sample.xlsx")

    feature_table = tables["Sample Table"]

    source = get_source_from_table(feature_table)
    properties = get_properties_from_table(feature_table)

    map = folium.Map()

    for property_name in properties:
        property = properties[property_name]

        feature_collection = ee.FeatureCollection(source)

        if property['TYPE'] == 'DESC' or property['TYPE'] == 'ID':
            continue
        if property['TYPE'] == "VALUE":
            values = get_values_from_table(feature_table,property_name)

            max_value = 0
            for name in values:
                value = values[name]
                if value is not None:
                    max_value = max(value, max_value)

            colors = []
            names = []
            for name in values:
                value = values[name]
                if value is None:
                    colors.append(rgba(125,125,125,125))
                else:
                    colors.append(red_hue(int(value / max_value * 255)))
                names.append(name)
            color_map = ee.Dictionary.fromLists(names,colors)
            def apply_color(feature):
                name = feature.get('name')
                feature = feature.set("style",{
                    "fillColor": color_map.get(name),
                })
                return feature

            feature_collection = feature_collection.map(apply_color).style(styleProperty="style", width=0.7, color="red")

        feature_collection_id = feature_collection.getMapId()
        layer = folium.TileLayer(
            tiles=feature_collection_id['tile_fetcher'].url_format,
            attr="Map Data",
            overlay=True,
            name= property_name
        )

        layer.add_to(map)

    map.add_child(folium.LayerControl())
    map.save('test.html')