# -*- coding: utf-8 -*-
"""
Created on Tue May 10 09:38:02 2022

@author: V1008647
"""
"""=====1======"""
import requests
import json

url = "https://10.224.81.70:6443/notify-service/api/notify"

a = "test mail mutiple 2"
b = "test phat an luon"
U_mail = "patrick.yy.tao@mail.foxconn.com"
from_mail = ""

payload = json.dumps({
    "toUser": U_mail,
    "toGroup": "",
    "from": from_mail,
    "message": "{\"title\": \"" + a + "\", \"body\": \"" + b + "\"}",
  
    "system": "MAIL",
    "source": "WS",
    "type": "TEXT"
})
headers = {
  'Content-Type': 'application/json'
}

response = requests.request("POST", url, headers=headers, data=payload, verify=False)
print(response.text)


