import requests
import json
import sys

KEY = sys.argv[len(sys.argv)-1]

if len(sys.argv) == 4:
    start_date = sys.argv[1]
    end_date = sys.argv[2]
    NEO_FEED_URL = f"https://api.nasa.gov/neo/rest/v1/feed?start_date={start_date}&end_date={end_date}&api_key={KEY}"
    x = requests.get(NEO_FEED_URL)
elif len(sys.argv) == 3:
    neo_id = sys.argv[1]
    NEO_LOOKUP = f"https://api.nasa.gov/neo/rest/v1/neo/{neo_id}?api_key={KEY}"
    x = requests.get(NEO_LOOKUP)
elif len(sys.argv) == 2:
    NEO_BROWSE = f"https://api.nasa.gov/neo/rest/v1/neo/browse?api_key={KEY}"
    x = requests.get(NEO_BROWSE)
else:
    print("Error in argument parsing")

if x.status_code == 200:
    try:                                    # if api call is successful, parse into a json named data
        data = x.json()
        print(data)
    except:                                 # otherwise print this
        print("JSON formatting failed")
else:                                       # if it is unsuccessful call, return the error code
    print(f"Returned error {x.status_code}")