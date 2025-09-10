import requests
import json

# 测试用户登录和数据访问
session = requests.Session()

print('=== 调试测试 ===')

# 1. 测试登录
print('\n1. 测试登录')
login_data = {
    'username': 'admin',
    'password': '123456'
}

login_response = session.post('http://localhost:5000/login', data=login_data, allow_redirects=False)
print(f'登录响应状态码: {login_response.status_code}')
print(f'登录响应头: {dict(login_response.headers)}')

if login_response.status_code == 302:
    print('✓ 登录成功，收到重定向')
    print(f'重定向到: {login_response.headers.get("Location")}')
else:
    print('✗ 登录失败或异常')
    print(f'响应内容: {login_response.text[:200]}...')

# 2. 检查session cookies
print('\n2. 检查session cookies')
print(f'Cookies: {dict(session.cookies)}')

# 3. 访问邮件数据页面
print('\n3. 访问邮件数据页面')
page_response = session.get('http://localhost:5000/mail_data')
print(f'页面访问状态码: {page_response.status_code}')

if page_response.status_code == 200:
    if 'mail_data.html' in page_response.url or 'mailDataTableBody' in page_response.text:
        print('✓ 成功访问邮件数据页面')
    else:
        print('✗ 可能被重定向到其他页面')
        print(f'当前URL: {page_response.url}')
else:
    print(f'✗ 页面访问失败: {page_response.status_code}')

# 4. 直接测试API
print('\n4. 测试API调用')
api_response = session.get('http://localhost:5000/api/mail_data?page=1&per_page=3')
print(f'API状态码: {api_response.status_code}')

try:
    api_data = api_response.json()
    print(f'API响应: {json.dumps(api_data, indent=2, ensure_ascii=False)}')
except Exception as e:
    print(f'API响应解析失败: {e}')
    print(f'原始响应: {api_response.text[:300]}...')

print('\n=== 测试完成 ===')