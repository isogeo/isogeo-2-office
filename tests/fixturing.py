# Standard library
import json
import logging
from os import environ, mkdir, path

# Isogeo
from isogeo_pysdk import Isogeo

# -- FIXTURES -----------------------------------------------------------------
share_id = environ.get('ISOGEO_API_DEV_ID')
share_token = environ.get('ISOGEO_API_DEV_SECRET')

# required dirs
if not path.exists("_logs"):
    mkdir(r"_logs")

if not path.exists("_auth"):
    mkdir(r"_auth")

# auth fixture
if not path.exists("_auth/client_secrets.json"):
    # fake dict
    auth_dict = {"web": {
                    "client_id": share_id,
                    "client_secret": share_token,
                    "auth_uri": "https://id.api.isogeo.com/oauth/authorize",
                    "token_uri": "https://id.api.isogeo.com/oauth/token"
                    }
                }

    # json dump
    with open("_auth/client_secrets.json", "w") as json_auth:
        json.dump(auth_dict,
                  json_auth)
    

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
