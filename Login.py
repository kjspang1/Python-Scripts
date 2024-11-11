# -*- coding: utf-8 -*-
"""
Created on Wed Mar  1 13:41:22 2023

@author: Kevin Spang
"""

def mysqlLogin():
    user = 'kevinspang'
    password = 'Mg$w%sEZ2sG4x7P'
    db = 'audience360'
    host = 'int-replica.addaptive.com'
    emailpw = 'America1'
    return(user, password, db, host, emailpw)


def snowflakeLogin():
    account ='eaa75578.us-east-1'
    user = 'scripts_kevin_spang'
    password = 'Addaptive1!'
    role = 'TEST_DB_KEVIN_SPANG'
    warehouse = 'KEVIN_SPANG'
    database = 'TEST_DB'
    schema = 'KEVIN_SPANG'
    passphrase = b'BqQkRbHWdpVdDblf24g5hvkRNPAlQpf3B4hRqj'
    private_key = b'''-----BEGIN ENCRYPTED PRIVATE KEY-----
MIIFHDBOBgkqhkiG9w0BBQ0wQTApBgkqhkiG9w0BBQwwHAQI+VLDAerdl48CAggA
MAwGCCqGSIb3DQIJBQAwFAYIKoZIhvcNAwcECMG73atBSC0lBIIEyFqN9sQt/tVO
pjnJJVNwSOz7CK9NEMbmi0zsCILkMIODwAp3+xlb1yZBrCVYE4duJlhkQHklkhph
zJ0uSgSCVY7sigDujmIlMD2qTUmk3G8pT5KgyA+UETmuv646dM1HrfoA5nQufY4d
MYYSPvap5e0KhLw7BXYP+M4qTlQjiZQJu1Jzw6XMbK9ZvdCeX2JYaEY+eMzXQa0y
uXrpX8gbfrXviuUV3o66B8tpAxpaPTf1xZkE2uheLscS6sTCoJgmQuZ1M3MDdQDo
0jhQzggALd/wulsZwZDvhDQnVBYC3pByN3TaingERS5bSGJHNvzkUNPoza8qz9WV
niRaN+rNrnCZooX4b/1D80FIhrFWMY6OYYPZghA8tG3nSwKOlwIDPuI6roUNS9/T
5eyPRN7F+ARZH7uwROQQCgkVmXZJaaOhuC2OPkf5aZmqFYRxd8iN7oVVXqjzzB1x
5Y4iZv6fB4LGknEPtjbjleqWnoyKTn7HxWlybYLpaszMnZqmHBxq7gNGbmo7dy2o
c/6Gn150MBZiCq15qnk1QfFez05UP73xlnPEM4i9ZBEoE1GjrhKnrmcKBNVlYeJp
NugYV2pLW0a9cG/IPjArVPnKKeAtnxADZztutu0i4BxIO4DZMGegINt9qeaeaD46
SocSJxS6mw44LwZ3Dc9UahOIgDzofcOphWDhXNRK8xFbuEIGOdPDJuBZhq0BLt3l
2alcdoWmh6ih4uGPjM94Lvhcvh9sDvXqC4WWdgA9bwmTgIyFDlhYvBzF6I9+ZczI
5r31olcso9nYtiXLwT93dyigG5ryIFGe+U8vP97TAYoYzYlV8Ah+jIzlCnE0BScv
1n/MXlkAvwVg+dQdB9zIbWQlBkdrZMRIcGdMUxTkETgQTaVWwBLeYldqQcDsH/vv
dzeV8M3b61idIvxqcO2N4+RFKQ07W5SVVCKkEk1X7ch2Ax8Vx6g5ovPOt8DF/08s
PDPFBGKg8mV7irBAiaoSCaUccruZZTl1ElVf8jm7b0UB3UyC3E6z9TWciqZrWsyt
xe44sVFDqgWnZKlLdGT/4iNMsSYi9iro07yZ1sV6ZUnm3k3g3CF3whqWdton6Wsf
sex49OYQoLQUvoehklK+CD0DmsI06w9Dxxpo7tVF7OfN7wRR+7bN3aavE2uzBfJB
C+TN45oppeqboGVtw07Dd5xpeJz8HzwGDliA5vOL6NZXRP3EYHC6Vz8iWCutsFYy
4doC2oOmIzyG3KEQ/dWDi/i9ksF/7TRm7DoFE+1JI7J17EGv+LzMV84iXvAyfGVk
QD+dxIGKfTyLu6CKKr6AtA1//13Qe/11ajO+5/85wwfZTkDkmqVNO2IDst/qKNRL
V8D+hLZkYsk44PAWXmZkT7ar0W+beErxIAO5iekQX5w6TpzdnzvcOe96nf41cBk0
Ri6r2bzBGWRUaXvv/kE9tddnFMwoRQxKoZnSZLfluvj/TAbO6wpY1MkFv7J1blj5
ySbE0ioG5i0CAmb1zG16refyFaHYbOHF9bwcyOEBOZH8tHrHck7Dn/ec7Im32+uc
QyS59qScNVBwLv4XZONBARJfz7hi3Qn7dCSC7kmVGQbDwX8QHUFnKDEyPhhBF0fZ
MkM6rAeIyXo5GTAPfkAfLA==
-----END ENCRYPTED PRIVATE KEY-----
                   '''
    return(account, user, password, role, warehouse, database, schema, passphrase, private_key)
 
 
 