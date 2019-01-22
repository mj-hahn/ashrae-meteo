import json
import mechanize
import requests
import argparse
import os

GOOGLE_MAPS_API_KEY = os.environ['GOOGLE_MAPS_API_KEY']
GOOGLE_MAPS_API_URL = 'https://maps.googleapis.com/maps/api/geocode/json'


def get_geocode(search_string):
    params = {
        'address': search_string,
        'key': GOOGLE_MAPS_API_KEY
    }
    req = requests.get(GOOGLE_MAPS_API_URL, params=params)
    res = req.json()
    return res.get("results", [])[0].get("geometry", {}).get("location", {})


def fetch_station(br, geocode):
    request_params = {
        "lat": geocode.get("lat"),
        "long": geocode.get("lng"),
        "number": "10",
        "ashrae_version": "2017"
    }
    url = "http://ashrae-meteo.info/request_places.php"
    new_request = mechanize.Request(
        url=url, data=request_params, method="POST")
    resp = br.open(new_request)
    resp_content = resp.read()
    resp_json = json.loads(resp_content)
    resp_str = json.dumps(resp_json, indent=4)

    stations = resp_json.get("meteo_stations", [])
    selected_station = stations[0]
    return selected_station


def fetch_weather_data(br, station_data):
    request_params = {
        "wmo": station_data.get("wmo"),
        "ashrae_version": "2017",
        "si_ip": "SI"
    }
    url = "http://ashrae-meteo.info/request_meteo_parametres.php"
    new_request = mechanize.Request(url=url, data=request_params, method="POST")
    resp = br.open(new_request)
    resp_content = resp.read()
    print(resp_content)


def main(args):

    # init mechanize
    br = mechanize.Browser()

    # get user input
    user_input = vars(args).get("location")[1]

    # get lat and long for search
    geocode = get_geocode(user_input)

    # get location data
    station = fetch_station(br, geocode)
    print(station)

    # get ashrae data
    fetch_weather_data(br, station)


if __name__ == "__main__":

    parser = argparse.ArgumentParser(description='Process some integers.')
    parser.add_argument('location', metavar='S', type=str, nargs='+', default='',
                        help='Address or latititude and longitude (xx.xx; yy.xx)')

    args = parser.parse_args()
    main(args)