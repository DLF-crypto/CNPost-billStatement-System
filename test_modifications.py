#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试账单信息管理修改的脚本
"""

import requests
import json

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

def test_bill_info_page(session):
    """测试账单信息管理页面"""
    print("\n=== 测试账单信息管理页面 ===")
    
    response = session.get(f'{BASE_URL}/bill_info')
    if response.status_code == 200:
        print("✓ 账单信息管理页面访问成功")
        # 检查页面内容是否包含正确的表头顺序
        content = response.text
        if '航班号</th>' in content and content.find('序号</th>') < content.find('航班号</th>'):
            print("✓ 表头顺序正确：航班号位于序号之后")
        else:
            print("✗ 表头顺序不正确")
        return True
    else:
        print("✗ 账单信息管理页面访问失败")
        return False

def test_add_bill_info(session):
    """测试添加账单信息功能"""
    print("\n=== 测试添加账单信息功能 ===")
    
    # 测试数据
    test_data1 = {
        'mail_class': 'TY',
        'des': 'ZRH',
        'route_info': '测试路由1',
        'flight_no': 'TEST001',
        'quote': '12.50',
        'carry_code': 'TESTCODE1'
    }
    
    test_data2 = {
        'mail_class': 'PY',
        'des': 'ZRH',  # 相同目的地，不同邮件类型，应该可以添加
        'route_info': '测试路由2',
        'flight_no': 'TEST002',
        'quote': '15.80',
        'carry_code': 'TESTCODE2'
    }
    
    test_data3 = {
        'mail_class': 'TY',
        'des': 'FRA',
        'route_info': '测试路由3',
        'flight_no': 'TEST001',  # 相同航班号，应该被拒绝
        'quote': '20.00',
        'carry_code': 'TESTCODE3'
    }
    
    # 添加第一条记录
    response = session.post(f'{BASE_URL}/add_bill_info', 
                           headers={'Content-Type': 'application/json'},
                           data=json.dumps(test_data1))
    
    if response.status_code == 200:
        result = response.json()
        if result.get('success'):
            print("✓ 第一条账单信息添加成功")
        else:
            print(f"✗ 第一条账单信息添加失败: {result.get('message')}")
    
    # 添加第二条记录（相同目的地，不同邮件类型）
    response = session.post(f'{BASE_URL}/add_bill_info', 
                           headers={'Content-Type': 'application/json'},
                           data=json.dumps(test_data2))
    
    if response.status_code == 200:
        result = response.json()
        if result.get('success'):
            print("✓ 相同目的地不同邮件类型的账单信息添加成功（验证取消唯一性限制）")
        else:
            print(f"✗ 相同目的地不同邮件类型的账单信息添加失败: {result.get('message')}")
    
    # 添加第三条记录（相同航班号）
    response = session.post(f'{BASE_URL}/add_bill_info', 
                           headers={'Content-Type': 'application/json'},
                           data=json.dumps(test_data3))
    
    if response.status_code == 200:
        result = response.json()
        if not result.get('success') and '航班号' in result.get('message', ''):
            print("✓ 相同航班号被正确拒绝（验证航班号唯一性）")
        else:
            print(f"✗ 相同航班号未被正确拒绝: {result.get('message')}")

def test_get_bill_info(session):
    """测试获取账单信息列表"""
    print("\n=== 测试获取账单信息列表 ===")
    
    response = session.get(f'{BASE_URL}/get_bill_info')
    if response.status_code == 200:
        result = response.json()
        if result.get('success'):
            print("✓ 获取账单信息列表成功")
            data = result.get('data', [])
            if data:
                print(f"  当前共有 {len(data)} 条账单信息")
                # 检查第一条记录的字段
                first_record = data[0]
                if all(field in first_record for field in ['mail_class', 'des', 'route_info', 'flight_no', 'quote', 'carry_code']):
                    print("✓ 账单信息字段完整")
                else:
                    print("✗ 账单信息字段不完整")
        else:
            print(f"✗ 获取账单信息列表失败: {result.get('message')}")
    else:
        print("✗ 获取账单信息列表请求失败")

def main():
    """主测试函数"""
    print("开始测试账单信息管理修改...")
    
    # 登录
    session = test_login()
    if not session:
        print("无法登录，测试终止")
        return
    
    # 测试各项功能
    test_bill_info_page(session)
    test_add_bill_info(session)
    test_get_bill_info(session)
    
    print("\n测试完成！")

if __name__ == '__main__':
    main()