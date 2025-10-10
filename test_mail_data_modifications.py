#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试邮件数据管理修改的脚本
"""

import requests
import json
import pandas as pd
import io

BASE_URL = 'http://127.0.0.1:5000'

def test_login():
    """测试登录功能"""
    print("=== 测试登录功能 ===")
    session = requests.Session()
    
    # 登录
    login_data = {
        'username': 'admin',
        'password': '123456'
    }
    
    response = session.post(f'{BASE_URL}/login', data=login_data)
    if response.status_code == 200 and 'dashboard' in response.url:
        print("✓ 登录成功")
        return session
    else:
        print("✗ 登录失败")
        return None

def test_mail_data_page(session):
    """测试邮件数据管理页面"""
    print("\n=== 测试邮件数据管理页面 ===")
    
    response = session.get(f'{BASE_URL}/mail_data')
    if response.status_code == 200:
        print("✓ 邮件数据管理页面访问成功")
        # 检查页面内容是否包含新的导入说明
        content = response.text
        if '只保留必填字段：总包号、接收时间、启运时间、到达时间、交邮时间' in content or '必填字段：总包号' in content:
            print("✓ 导入说明已更新")
        else:
            print("✗ 导入说明未正确更新")
        
        if '是否有航班号' in content:
            print("✓ 航班号筛选选项已添加")
        else:
            print("✗ 航班号筛选选项未添加")
        return True
    else:
        print("✗ 邮件数据管理页面访问失败")
        return False

def test_flight_no_api(session):
    """测试根据航班号获取账单信息API"""
    print("\n=== 测试航班号查询API ===")
    
    # 测试根据航班号查询账单信息
    test_data = {
        'flight_no': 'TEST001'  # 假设这是之前添加的测试航班号
    }
    
    response = session.post(f'{BASE_URL}/api/get_bill_info_by_flight_no', 
                           headers={'Content-Type': 'application/json'},
                           data=json.dumps(test_data))
    
    if response.status_code == 200:
        result = response.json()
        if result.get('success'):
            print("✓ 根据航班号查询账单信息成功")
            print(f"  查询结果：{result.get('data')}")
        else:
            print(f"✗ 根据航班号查询失败: {result.get('message')}")
    else:
        print("✗ 航班号查询API请求失败")

def test_mail_data_search(session):
    """测试邮件数据搜索功能"""
    print("\n=== 测试邮件数据搜索功能 ===")
    
    # 测试航班号筛选
    search_params = {
        'has_flight_no': 'yes'
    }
    
    params = {
        'page': 1,
        'per_page': 10,
        'search': json.dumps(search_params)
    }
    
    response = session.get(f'{BASE_URL}/api/mail_data', params=params)
    if response.status_code == 200:
        result = response.json()
        if result.get('success'):
            print("✓ 航班号筛选功能正常")
            data = result.get('data', [])
            print(f"  找到 {len(data)} 条有航班号的记录")
        else:
            print(f"✗ 航班号筛选失败: {result.get('message')}")
    else:
        print("✗ 邮件数据搜索请求失败")

def create_test_excel():
    """创建测试用的Excel文件"""
    print("\n=== 创建测试Excel文件 ===")
    
    # 创建测试数据
    test_data = [
        {
            '总包号': 'ABCDEFGHIJKLMNO12345678901234',  # 29位：15位字母 + 14位数字
            '接收时间': '2025-01-15 10:00:00',
            '启运时间': '2025-01-15 12:00:00',
            '到达时间': '2025-01-16 08:00:00',
            '交邮时间': '2025-01-16 10:00:00',
            '航班号': 'TEST002'
        },
        {
            '总包号': 'ZYXWVUTSRQPONML98765432109876',  # 29位：15位字母 + 14位数字
            '接收时间': '2025-01-15 11:00:00',
            '启运时间': '2025-01-15 13:00:00',
            '到达时间': '2025-01-16 09:00:00',
            '交邮时间': '2025-01-16 11:00:00'
            # 不包含航班号，测试空值处理
        }
    ]
    
    df = pd.DataFrame(test_data)
    
    # 保存为Excel文件
    excel_buffer = io.BytesIO()
    df.to_excel(excel_buffer, index=False, engine='openpyxl')
    excel_buffer.seek(0)
    
    print("✓ 测试Excel文件创建成功")
    return excel_buffer.getvalue()

def test_excel_import(session):
    """测试Excel导入功能"""
    print("\n=== 测试Excel导入功能 ===")
    
    # 创建测试Excel文件
    excel_data = create_test_excel()
    
    # 准备文件上传
    files = {
        'excel_file': ('test_mail_data.xlsx', excel_data, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    }
    
    response = session.post(f'{BASE_URL}/api/import_excel_mail_data', files=files)
    
    if response.status_code == 200:
        result = response.json()
        if result.get('success'):
            print("✓ Excel导入成功")
            print(f"  导入数量：{result.get('imported_count', 0)}")
        else:
            print(f"✗ Excel导入失败: {result.get('message')}")
            # 如果有错误详情，显示出来
            if 'errors' in result:
                print("  错误详情：")
                for error in result['errors'][:3]:  # 只显示前3个错误
                    print(f"    - 第{error['row']}行: {error['message']}")
    else:
        print("✗ Excel导入请求失败")

def main():
    """主测试函数"""
    print("开始测试邮件数据管理修改...")
    
    # 登录
    session = test_login()
    if not session:
        print("无法登录，测试终止")
        return
    
    # 测试各项功能
    test_mail_data_page(session)
    test_flight_no_api(session)
    test_mail_data_search(session)
    test_excel_import(session)
    
    print("\n测试完成！")

if __name__ == '__main__':
    main()