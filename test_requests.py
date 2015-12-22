# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (print_function, unicode_literals)


import requests
# import dateutil
import base64


idu = "plugin-qgis-dev-8f82da32ffb643d9b039fe44b4e8c62d"
secret = "SCTezaDWBvJO1fnsenB1ZHWWZzNBB8wkRw2727Qvt72M41dzFVQSxEUTBJcMiK0I"

to64 = idu + ':' + secret

toheader = base64.b64encode(to64)

headers = {'Authorization': 'Basic ' + toheader}
payload = {'grant_type':'client_credentials'}

p = requests.post('https://id.api.isogeo.com/oauth/token', headers=headers, data=payload)

# print(p.text)
# print(dir(p))

axx = p.json()
print(axx.keys())

head = {'Authorization': 'Bearer ' + axx.get("access_token")}
r = requests.get('https://v1.api.isogeo.com/resources/search', headers=head)


# print(r.json())



