# Standard library
import json
import logging
from os import environ, path

# Isogeo
from isogeo_pysdk import Isogeo

# -- FIXTURES -----------------------------------------------------------------
share_id = environ.get('ISOGEO_API_DEV_ID')
share_token = environ.get('ISOGEO_API_DEV_SECRET')

# instanciating the class
isogeo = Isogeo(client_id=share_id,
                client_secret=share_token)
token = isogeo.connect()

# Downloading directly from Isogeo API
BASE_DIR = path.dirname(path.abspath(__file__))

# complete search - only Isogeo Tests
out_search_complete_tests =  path.join(BASE_DIR,
                                       "fixtures",
                                       "api_search_complete_tests.json")
if not path.isfile(out_search_complete_tests):
    request = isogeo.search(token, query="owner:32f7e95ec4e94ca3bc1afda960003882",
                            whole_share=1, include="all", augment=1)
    with open(out_search_complete_tests, "w") as json_basic:
        json.dump(request,
                json_basic,
                sort_keys=True
                )
else:
    logging.info("JSON already exists. If you want to update it, delete it first.")

# complete search
out_search_complete = path.join(BASE_DIR,
                                "fixtures",
                                "api_search_complete.json")
if not path.isfile(out_search_complete):
request = isogeo.search(token,
                        whole_share=1, include="all", augment=1)
with open(path.join(BASE_DIR, "fixtures", "api_search_complete.json"), "w") as json_basic:
    json.dump(request,
              json_basic,
              sort_keys=True,
              )
else:
    logging.info(
        "JSON already exists. If you want to update it, delete it first.")
