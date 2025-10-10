import requests
import json

# 登录
session = requests.Session()
session.post('http://127.0.0.1:5000/login', data={'username': 'admin', 'password': '123456'})

# 查询有航班号的记录
response = session.get('http://127.0.0.1:5000/api/mail_data', params={
    'page': 1, 
    'per_page': 5, 
    'search': json.dumps({'has_flight_no': 'yes'})
})
result = response.json()
print(f'有航班号的记录数: {len(result.get("data", []))}')

# 查询无航班号的记录
response = session.get('http://127.0.0.1:5000/api/mail_data', params={
    'page': 1, 
    'per_page': 5, 
    'search': json.dumps({'has_flight_no': 'no'})
})
result = response.json()
print(f'无航班号的记录数: {len(result.get("data", []))}')

print('验证完成！')