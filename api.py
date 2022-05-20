import requests

res = requests.get("https://www.bcbsnm.com/community-centennial/json/cc-provider-directory-nm.json")

print(res.json())