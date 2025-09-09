from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
import mysql.connector
from werkzeug.security import check_password_hash, generate_password_hash
from datetime import datetime
import pandas as pd
from werkzeug.utils import secure_filename
import os
from mysql.connector import Error
import json
from datetime import datetime
import re

app = Flask(__name__)

# 配置密钥
app.secret_key = 'cnpost_invoice_system_2024'

# 配置上传文件夹
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# 确保上传文件夹存在
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# 字段验证函数
def validate_mail_class(mail_class):
    """验证邮件类型"""
    if not mail_class:
        return False, "邮件类型不能为空"
    if mail_class not in ['TY', 'PY']:
        return False, "邮件类型字段不符合要求"
    return True, ""

def validate_destination(destination):
    """验证目的地格式"""
    if not destination:
        return False, "目的地不能为空"
    if not re.match(r'^[A-Z]{3}$', destination):
        return False, "目的地字段不符合要求"
    return True, ""

def validate_quote(quote):
    """验证报价格式"""
    if quote is None or quote == '':
        return False, "报价不能为空"
    try:
        quote_float = float(quote)
        if quote_float < 0:
            return False, "报价字段不符合要求"
        # 检查小数位数是否超过2位
        if len(str(quote_float).split('.')[-1]) > 2 and '.' in str(quote_float):
            return False, "报价字段不符合要求"
        return True, ""
    except (ValueError, TypeError):
        return False, "报价字段不符合要求"

def check_destination_uniqueness(mail_class, destination, connection, exclude_id=None):
    """检查目的地唯一性"""
    try:
        cursor = connection.cursor()
        if exclude_id:
            query = "SELECT COUNT(*) FROM bill_info WHERE mail_class = %s AND des = %s AND id != %s"
            cursor.execute(query, (mail_class, destination, exclude_id))
        else:
            query = "SELECT COUNT(*) FROM bill_info WHERE mail_class = %s AND des = %s"
            cursor.execute(query, (mail_class, destination))
        
        count = cursor.fetchone()[0]
        cursor.close()
        
        if count > 0:
            return False, "该目的地已有报价"
        return True, ""
    except Exception as e:
        return False, f"验证失败: {str(e)}"

def validate_bill_data(data, connection, exclude_id=None):
    """验证账单数据"""
    errors = []
    
    # 验证邮件类型
    is_valid, error_msg = validate_mail_class(data.get('mail_class', ''))
    if not is_valid:
        errors.append(error_msg)
    
    # 验证目的地格式
    is_valid, error_msg = validate_destination(data.get('des', ''))
    if not is_valid:
        errors.append(error_msg)
    
    # 验证报价
    is_valid, error_msg = validate_quote(data.get('quote'))
    if not is_valid:
        errors.append(error_msg)
    
    # 验证路由信息（必填）
    route_info = data.get('route_info', '').strip()
    if not route_info:
        errors.append('路由信息不能为空')
    
    # 验证航班号（必填）
    flight_no = data.get('flight_no', '').strip()
    if not flight_no:
        errors.append('航班号不能为空')
    
    # 验证承运代码（必填）
    carry_code = data.get('carry_code', '').strip()
    if not carry_code:
        errors.append('承运代码不能为空')
    
    # 检查目的地唯一性（只要邮件类型和目的地格式正确就检查）
    mail_class = data.get('mail_class', '')
    destination = data.get('des', '')
    
    # 只有当邮件类型和目的地都有效时才检查唯一性
    if mail_class in ['TY', 'PY'] and destination and len(destination) == 3 and destination.isalpha():
        is_valid, error_msg = check_destination_uniqueness(
            mail_class, 
            destination.upper(), 
            connection, 
            exclude_id
        )
        if not is_valid:
            errors.append(error_msg)
    
    return len(errors) == 0, errors

# MySQL数据库配置
DB_CONFIG = {
    'host': 'localhost',
    'database': 'cnpost_bill_system',
    'user': 'root',
    'password': '123456',
    'charset': 'utf8mb4'
}

# 数据库连接函数
def get_db_connection():
    """获取数据库连接"""
    try:
        # 先连接到MySQL服务器（不指定数据库）
        server_config = DB_CONFIG.copy()
        if 'database' in server_config:
            del server_config['database']
        
        connection = mysql.connector.connect(**server_config)
        cursor = connection.cursor()
        
        # 创建数据库（如果不存在）
        cursor.execute("CREATE DATABASE IF NOT EXISTS cnpost_bill_system CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
        cursor.close()
        connection.close()
        
        # 重新连接到指定数据库
        full_config = DB_CONFIG.copy()
        full_config['database'] = 'cnpost_bill_system'
        connection = mysql.connector.connect(**full_config)
        return connection
    except Error as e:
        print(f"数据库连接错误: {e}")
        return None

# 初始化数据库表
def init_database():
    connection = get_db_connection()
    if connection:
        try:
            cursor = connection.cursor()
            
            # 创建员工表
            create_employees_table = """
            CREATE TABLE IF NOT EXISTS employees (
                id INT AUTO_INCREMENT PRIMARY KEY,
                employee_id VARCHAR(20) UNIQUE NOT NULL,
                username VARCHAR(50) UNIQUE NOT NULL,
                password VARCHAR(255) NOT NULL,
                name VARCHAR(100) NOT NULL,
                phone VARCHAR(20),
                email VARCHAR(100),
                role_type ENUM('管理员', '一般员工') DEFAULT '一般员工',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
            )
            """
            
            cursor.execute(create_employees_table)
            
            # 检查是否已存在admin用户
            cursor.execute("SELECT COUNT(*) FROM employees WHERE username = 'admin'")
            admin_exists = cursor.fetchone()[0]
            
            if admin_exists == 0:
                # 插入默认admin用户
                insert_admin = """
                INSERT INTO employees (employee_id, username, password, name, phone, email, role_type)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                """
                admin_data = (
                    'YC001',
                    'admin',
                    generate_password_hash('123456'),
                    '系统管理员',
                    '13800138000',
                    'admin@cnpost.com',
                    '管理员'
                )
                cursor.execute(insert_admin, admin_data)
            
            connection.commit()
            print("数据库初始化成功")
            
        except Error as e:
            print(f"数据库初始化错误: {e}")
        finally:
            cursor.close()
            connection.close()

# 验证用户登录
def verify_user(username, password):
    connection = get_db_connection()
    if connection:
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute("SELECT * FROM employees WHERE username = %s", (username,))
            user = cursor.fetchone()
            
            if user and check_password_hash(user['password'], password):
                return user
            return None
            
        except Error as e:
            print(f"用户验证错误: {e}")
            return None
        finally:
            cursor.close()
            connection.close()
    return None

def get_user_permissions(user_id):
    """获取用户的角色权限"""
    try:
        connection = get_db_connection()
        cursor = connection.cursor(dictionary=True)
        
        # 查询用户的角色权限
        query = """
        SELECT r.permissions 
        FROM employees e 
        JOIN roles r ON e.role_id = r.id 
        WHERE e.id = %s
        """
        cursor.execute(query, (user_id,))
        result = cursor.fetchone()
        
        if result and result['permissions']:
            return result['permissions'].split(',')
        return []
        
    except mysql.connector.Error as err:
        print(f"获取用户权限时出错: {err}")
        return []
    finally:
        cursor.close()
        connection.close()

@app.route('/')
def index():
    """首页路由，重定向到登录页面"""
    if 'loggedin' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    """登录页面"""
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        
        # 使用数据库验证用户
        user = verify_user(username, password)
        if user:
            session['loggedin'] = True
            session['username'] = user['username']
            session['user_id'] = user['id']
            session['user_name'] = user['name']
            session['role_type'] = user['role_type']
            flash('登录成功！', 'success')
            return redirect(url_for('dashboard'))
        else:
            flash('用户名或密码错误！', 'error')
    
    return render_template('login.html')

@app.route('/dashboard')
def dashboard():
    """登录后的主页面"""
    if 'loggedin' not in session:
        return redirect(url_for('login'))
    
    # 获取用户权限
    user_permissions = get_user_permissions(session['user_id'])
    
    return render_template('dashboard.html', 
                         username=session.get('user_name', session['username']),
                         permissions=user_permissions)

# 员工管理页面
@app.route('/employees')
def employees():
    if 'loggedin' not in session:
        return redirect(url_for('login'))
    
    # 检查用户是否有员工管理权限
    user_permissions = get_user_permissions(session['user_id'])
    if 'employees' not in user_permissions:
        flash('您没有访问员工管理的权限！', 'error')
        return redirect(url_for('dashboard'))
    
    employees_list = get_all_employees()
    roles_list = get_all_roles()  # 获取所有角色用于下拉选择
    return render_template('employees.html', 
                         username=session.get('user_name', session['username']),
                         employees=employees_list,
                         roles=roles_list,
                         permissions=user_permissions)

# 角色管理页面
@app.route('/roles')
def roles():
    if 'loggedin' not in session:
        return redirect(url_for('login'))
    
    # 检查用户是否有角色管理权限
    user_permissions = get_user_permissions(session['user_id'])
    if 'roles' not in user_permissions:
        flash('您没有访问角色管理的权限！', 'error')
        return redirect(url_for('dashboard'))
    
    roles_list = get_all_roles()
    return render_template('roles.html', 
                         username=session.get('user_name', session['username']),
                         roles=roles_list,
                         permissions=user_permissions)

# 产品信息管理页面
@app.route('/products')
def products():
    if 'loggedin' not in session:
        return redirect(url_for('login'))
    
    # 检查用户是否有产品管理权限
    user_permissions = get_user_permissions(session['user_id'])
    if 'products' not in user_permissions:
        flash('您没有访问产品信息管理的权限！', 'error')
        return redirect(url_for('dashboard'))
    
    products_list = get_all_products()
    return render_template('products.html', 
                         username=session.get('user_name', session['username']),
                         products=products_list,
                         permissions=user_permissions)

# 邮件数据管理页面
@app.route('/mail_data')
def mail_data():
    if 'loggedin' not in session:
        return redirect(url_for('login'))
    
    # 检查用户是否有邮件数据管理权限
    user_permissions = get_user_permissions(session['user_id'])
    if 'mail_data' not in user_permissions:
        flash('您没有访问邮件数据管理的权限！', 'error')
        return redirect(url_for('dashboard'))
    
    return render_template('mail_data.html', 
                         username=session.get('user_name', session['username']),
                         permissions=user_permissions)

# 获取所有角色
def get_all_roles():
    """获取所有角色信息"""
    connection = get_db_connection()
    if connection is None:
        return []
    
    try:
        cursor = connection.cursor(dictionary=True)
        cursor.execute("""
            SELECT id, role_name, permissions, created_at, updated_at 
            FROM roles 
            ORDER BY id
        """)
        roles = cursor.fetchall()
        return roles
    except Error as e:
        print(f"获取角色列表错误: {e}")
        return []
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()

# 获取所有产品
def get_all_products():
    """获取所有产品信息"""
    connection = get_db_connection()
    if connection is None:
        return []
    
    try:
        cursor = connection.cursor(dictionary=True)
        cursor.execute("""
            SELECT product_id, product_identifier1, product_identifier2, 
                   product_settle_code, product_name, created_at, updated_at 
            FROM products 
            ORDER BY created_at DESC
        """)
        products = cursor.fetchall()
        return products
    except Error as e:
        print(f"获取产品列表错误: {e}")
        return []
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()

# 获取所有员工
def get_all_employees():
    connection = get_db_connection()
    if connection:
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute("""
                SELECT e.*, r.role_name as role_type 
                FROM employees e 
                LEFT JOIN roles r ON e.role_id = r.id 
                ORDER BY e.created_at DESC
            """)
            employees = cursor.fetchall()
            return employees
        except Error as e:
            print(f"获取员工列表错误: {e}")
            return []
        finally:
            cursor.close()
            connection.close()
    return []

# 生成新的员工ID
def generate_employee_id():
    connection = get_db_connection()
    if connection:
        try:
            cursor = connection.cursor()
            cursor.execute("SELECT employee_id FROM employees ORDER BY employee_id DESC LIMIT 1")
            result = cursor.fetchone()
            
            if result:
                last_id = result[0]
                # 提取数字部分并加1
                num = int(last_id[2:]) + 1
                return f"YC{num:03d}"
            else:
                return "YC001"
                
        except Error as e:
            print(f"生成员工ID错误: {e}")
            return "YC001"
        finally:
            cursor.close()
            connection.close()
    return "YC001"

# 添加员工
@app.route('/add_employee', methods=['POST'])
def add_employee():
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        data = request.get_json()
        employee_id = generate_employee_id()
        
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            
            # 检查用户名是否已存在
            cursor.execute("SELECT COUNT(*) FROM employees WHERE username = %s", (data['username'],))
            if cursor.fetchone()[0] > 0:
                return jsonify({'success': False, 'message': '用户名已存在'})
            
            # 根据角色名称获取role_id
            role_name = data.get('role_type', '普通用户')
            cursor.execute("SELECT id FROM roles WHERE role_name = %s", (role_name,))
            role_result = cursor.fetchone()
            role_id = role_result[0] if role_result else 2  # 默认为普通用户角色ID
            
            # 插入新员工
            insert_query = """
            INSERT INTO employees (employee_id, username, password, name, phone, email, role_id)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            """
            
            employee_data = (
                employee_id,
                data['username'],
                generate_password_hash(data.get('password', '123456')),
                data['name'],
                data.get('phone', ''),
                data.get('email', ''),
                role_id
            )
            
            cursor.execute(insert_query, employee_data)
            connection.commit()
            
            return jsonify({'success': True, 'message': '员工添加成功', 'employee_id': employee_id})
            
    except Error as e:
        return jsonify({'success': False, 'message': f'添加失败: {str(e)}'})
    finally:
        if connection:
            if 'cursor' in locals():
                cursor.close()
            connection.close()

# 更新员工
@app.route('/update_employee', methods=['POST'])
def update_employee():
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        data = request.get_json()
        employee_id = data['employee_id']
        
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            
            # 检查用户名是否被其他员工使用
            cursor.execute("SELECT COUNT(*) FROM employees WHERE username = %s AND employee_id != %s", 
                         (data['username'], employee_id))
            if cursor.fetchone()[0] > 0:
                return jsonify({'success': False, 'message': '用户名已被其他员工使用'})
            
            # 根据角色名称获取role_id
            role_name = data.get('role_type', '普通用户')
            cursor.execute("SELECT id FROM roles WHERE role_name = %s", (role_name,))
            role_result = cursor.fetchone()
            role_id = role_result[0] if role_result else 2  # 默认为普通用户角色ID
            
            # 更新员工信息
            update_query = """
            UPDATE employees 
            SET username = %s, name = %s, phone = %s, email = %s, role_id = %s
            WHERE employee_id = %s
            """
            
            employee_data = (
                data['username'],
                data['name'],
                data.get('phone', ''),
                data.get('email', ''),
                role_id,
                employee_id
            )
            
            cursor.execute(update_query, employee_data)
            connection.commit()
            
            return jsonify({'success': True, 'message': '员工信息更新成功'})
            
    except Error as e:
        return jsonify({'success': False, 'message': f'更新失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 删除员工
@app.route('/delete_employee', methods=['POST'])
def delete_employee():
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        data = request.get_json()
        employee_id = data['employee_id']
        
        # 不允许删除YC001（admin用户）
        if employee_id == 'YC001':
            return jsonify({'success': False, 'message': '不能删除系统管理员账户'})
        
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            cursor.execute("DELETE FROM employees WHERE employee_id = %s", (employee_id,))
            connection.commit()
            
            return jsonify({'success': True, 'message': '员工删除成功'})
            
    except Error as e:
        return jsonify({'success': False, 'message': f'删除失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 添加角色
@app.route('/add_role', methods=['POST'])
def add_role():
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        data = request.get_json()
        role_name = data['role_name'].strip()
        permissions = ','.join(data.get('permissions', []))
        
        if not role_name:
            return jsonify({'success': False, 'message': '角色名称不能为空'})
        
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            cursor.execute(
                "INSERT INTO roles (role_name, permissions) VALUES (%s, %s)",
                (role_name, permissions)
            )
            connection.commit()
            return jsonify({'success': True, 'message': '角色添加成功'})
            
    except Error as e:
        if 'Duplicate entry' in str(e):
            return jsonify({'success': False, 'message': '角色名称已存在'})
        return jsonify({'success': False, 'message': f'添加失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 更新角色
@app.route('/update_role', methods=['POST'])
def update_role():
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        data = request.get_json()
        role_id = data['role_id']
        role_name = data['role_name'].strip()
        permissions = ','.join(data.get('permissions', []))
        
        if not role_name:
            return jsonify({'success': False, 'message': '角色名称不能为空'})
        
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            cursor.execute(
                "UPDATE roles SET role_name = %s, permissions = %s WHERE id = %s",
                (role_name, permissions, role_id)
            )
            connection.commit()
            return jsonify({'success': True, 'message': '角色更新成功'})
            
    except Error as e:
        if 'Duplicate entry' in str(e):
            return jsonify({'success': False, 'message': '角色名称已存在'})
        return jsonify({'success': False, 'message': f'更新失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 删除角色
@app.route('/delete_role', methods=['POST'])
def delete_role():
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        data = request.get_json()
        role_id = data['role_id']
        
        # 检查是否有员工使用此角色
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            cursor.execute("SELECT COUNT(*) FROM employees WHERE role_id = %s", (role_id,))
            count = cursor.fetchone()[0]
            
            if count > 0:
                return jsonify({'success': False, 'message': f'无法删除，还有 {count} 个员工使用此角色'})
            
            cursor.execute("DELETE FROM roles WHERE id = %s", (role_id,))
            connection.commit()
            return jsonify({'success': True, 'message': '角色删除成功'})
            
    except Error as e:
        return jsonify({'success': False, 'message': f'删除失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 新增产品
@app.route('/add_product', methods=['POST'])
def add_product():
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        product_identifier1 = request.form.get('product_identifier1')
        product_identifier2 = request.form.get('product_identifier2')
        product_settle_code = request.form.get('product_settle_code')
        product_name = request.form.get('product_name')
        
        if not all([product_identifier1, product_identifier2, product_settle_code, product_name]):
            return jsonify({'success': False, 'message': '请填写所有必填字段'})
        
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            cursor.execute("""
                INSERT INTO products (product_identifier1, product_identifier2, product_settle_code, product_name, created_at)
                VALUES (%s, %s, %s, %s, NOW())
            """, (product_identifier1, product_identifier2, product_settle_code, product_name))
            connection.commit()
            return jsonify({'success': True, 'message': '产品添加成功'})
            
    except Error as e:
        if 'Duplicate entry' in str(e):
            return jsonify({'success': False, 'message': '产品信息已存在'})
        return jsonify({'success': False, 'message': f'添加失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 获取单个产品信息
@app.route('/get_product/<int:product_id>')
def get_product(product_id):
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor(dictionary=True)
            cursor.execute("""
                SELECT product_id, product_identifier1, product_identifier2, 
                       product_settle_code, product_name, created_at, updated_at
                FROM products WHERE product_id = %s
            """, (product_id,))
            product = cursor.fetchone()
            
            if product:
                return jsonify({'success': True, 'product': product})
            else:
                return jsonify({'success': False, 'message': '产品不存在'})
                
    except Error as e:
        return jsonify({'success': False, 'message': f'获取产品信息失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 更新产品
@app.route('/update_product/<int:product_id>', methods=['PUT'])
def update_product(product_id):
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        product_identifier1 = request.form.get('product_identifier1')
        product_identifier2 = request.form.get('product_identifier2')
        product_settle_code = request.form.get('product_settle_code')
        product_name = request.form.get('product_name')
        
        if not all([product_identifier1, product_identifier2, product_settle_code, product_name]):
            return jsonify({'success': False, 'message': '请填写所有必填字段'})
        
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            cursor.execute("""
                UPDATE products 
                SET product_identifier1 = %s, product_identifier2 = %s, product_settle_code = %s, 
                    product_name = %s, updated_at = NOW()
                WHERE product_id = %s
            """, (product_identifier1, product_identifier2, product_settle_code, product_name, product_id))
            connection.commit()
            return jsonify({'success': True, 'message': '产品更新成功'})
            
    except Error as e:
        if 'Duplicate entry' in str(e):
            return jsonify({'success': False, 'message': '产品信息已存在'})
        return jsonify({'success': False, 'message': f'更新失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 删除产品
@app.route('/delete_product/<int:product_id>', methods=['DELETE'])
def delete_product(product_id):
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            cursor.execute("DELETE FROM products WHERE product_id = %s", (product_id,))
            connection.commit()
            return jsonify({'success': True, 'message': '产品删除成功'})
            
    except Error as e:
        return jsonify({'success': False, 'message': f'删除失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 检查产品识别符1是否存在
@app.route('/check_product_identifier1/<identifier1>', methods=['GET'])
def check_product_identifier1(identifier1):
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            cursor.execute("SELECT COUNT(*) FROM products WHERE product_identifier1 = %s", (identifier1,))
            count = cursor.fetchone()[0]
            return jsonify({'exists': count > 0})
            
    except Error as e:
        return jsonify({'success': False, 'message': f'查询失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 检查产品识别符组合是否存在
@app.route('/check_product_identifiers/<identifier1>/<identifier2>', methods=['GET'])
def check_product_identifiers(identifier1, identifier2):
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            cursor.execute("SELECT COUNT(*) FROM products WHERE product_identifier1 = %s AND product_identifier2 = %s", (identifier1, identifier2))
            count = cursor.fetchone()[0]
            return jsonify({'exists': count > 0})
            
    except Error as e:
        return jsonify({'success': False, 'message': f'查询失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 检查产品识别符1是否存在（排除指定产品）
@app.route('/check_product_identifier1_exclude/<identifier1>/<int:product_id>', methods=['GET'])
def check_product_identifier1_exclude(identifier1, product_id):
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            cursor.execute("SELECT COUNT(*) FROM products WHERE product_identifier1 = %s AND product_id != %s", (identifier1, product_id))
            count = cursor.fetchone()[0]
            return jsonify({'exists': count > 0})
            
    except Error as e:
        return jsonify({'success': False, 'message': f'查询失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 检查产品识别符组合是否存在（排除指定产品）
@app.route('/check_product_identifiers_exclude/<identifier1>/<identifier2>/<int:product_id>', methods=['GET'])
def check_product_identifiers_exclude(identifier1, identifier2, product_id):
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            cursor.execute("SELECT COUNT(*) FROM products WHERE product_identifier1 = %s AND product_identifier2 = %s AND product_id != %s", (identifier1, identifier2, product_id))
            count = cursor.fetchone()[0]
            return jsonify({'exists': count > 0})
            
    except Error as e:
        return jsonify({'success': False, 'message': f'查询失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 获取所有产品数据（用于分页显示）
@app.route('/get_products', methods=['GET'])
def get_products():
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor(dictionary=True)
            cursor.execute("SELECT * FROM products ORDER BY created_at DESC")
            products = cursor.fetchall()
            
            # 转换日期格式
            for product in products:
                if product['created_at']:
                    product['created_at'] = product['created_at'].isoformat()
            
            return jsonify({'success': True, 'products': products})
            
    except Error as e:
        return jsonify({'success': False, 'message': f'获取产品数据失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 导入产品Excel
@app.route('/import_products', methods=['POST'])
def import_products():
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': '没有选择文件'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'message': '没有选择文件'})
    
    if file and allowed_file(file.filename):
        try:
            # 读取Excel文件
            df = pd.read_excel(file)
            
            # 清理列名（去除空格和特殊字符）
            df.columns = df.columns.str.strip()
            
            # 检查必需的列
            required_columns = ['产品识别符1', '产品结算代码', '产品中文名称']
            actual_columns = list(df.columns)
            
            # 调试信息：记录实际的列名
            print(f"Excel文件实际列名: {actual_columns}")
            
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                return jsonify({
                    'success': False, 
                    'message': f'Excel文件缺少必需的列: {", ".join(missing_columns)}。实际列名: {", ".join(actual_columns)}'
                })
            
            connection = get_db_connection()
            if not connection:
                return jsonify({'success': False, 'message': '数据库连接失败'})
            
            cursor = connection.cursor()
            success_count = 0
            error_messages = []
            
            for index, row in df.iterrows():
                try:
                    # 获取数据
                    identifier1 = str(row['产品识别符1']).strip() if pd.notna(row['产品识别符1']) else ''
                    identifier2 = str(row['产品识别符2']).strip() if '产品识别符2' in df.columns and pd.notna(row['产品识别符2']) else ''
                    settle_code = str(row['产品结算代码']).strip() if pd.notna(row['产品结算代码']) else ''
                    product_name = str(row['产品中文名称']).strip() if pd.notna(row['产品中文名称']) else ''
                    
                    # 验证必填项
                    if not identifier1 or not settle_code or not product_name:
                        error_messages.append(f'第{index+2}行: 产品识别符1、产品结算代码、产品中文名称为必填项')
                        continue
                    
                    # 检查产品识别符1是否重复
                    cursor.execute("SELECT COUNT(*) FROM products WHERE product_identifier1 = %s", (identifier1,))
                    if cursor.fetchone()[0] > 0:
                        if not identifier2:
                            error_messages.append(f'第{index+2}行: 产品识别符1已存在，产品识别符2不能为空')
                            continue
                        # 检查组合是否重复
                        cursor.execute("SELECT COUNT(*) FROM products WHERE product_identifier1 = %s AND product_identifier2 = %s", (identifier1, identifier2))
                        if cursor.fetchone()[0] > 0:
                            error_messages.append(f'第{index+2}行: 相同产品识别符1下的产品识别符2已存在')
                            continue
                    
                    # 插入数据
                    insert_query = "INSERT INTO products (product_identifier1, product_identifier2, product_settle_code, product_name) VALUES (%s, %s, %s, %s)"
                    cursor.execute(insert_query, (identifier1, identifier2, settle_code, product_name))
                    success_count += 1
                    
                except Exception as e:
                    error_messages.append(f'第{index+2}行: {str(e)}')
                    continue
            
            connection.commit()
            
            if error_messages:
                return jsonify({
                    'success': True, 
                    'count': success_count,
                    'message': f'部分导入成功，共导入{success_count}条记录。错误信息：' + '; '.join(error_messages[:5])  # 只显示前5个错误
                })
            else:
                return jsonify({'success': True, 'count': success_count})
                
        except Exception as e:
            return jsonify({'success': False, 'message': f'导入失败: {str(e)}'})
        finally:
            if 'cursor' in locals() and cursor:
                cursor.close()
            if 'connection' in locals() and connection:
                connection.close()
    else:
        return jsonify({'success': False, 'message': '不支持的文件格式，请上传.xlsx或.xls文件'})

# 重置员工密码
@app.route('/reset_password', methods=['POST'])
def reset_password():
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        data = request.get_json()
        employee_id = data['employee_id']
        new_password = data.get('new_password', '123456')
        
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            
            # 更新密码
            update_query = "UPDATE employees SET password = %s WHERE employee_id = %s"
            cursor.execute(update_query, (generate_password_hash(new_password), employee_id))
            connection.commit()
            
            return jsonify({'success': True, 'message': '密码重置成功'})
            
    except Error as e:
        return jsonify({'success': False, 'message': f'重置失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 修改密码
@app.route('/change_password', methods=['POST'])
def change_password():
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        data = request.get_json()
        current_password = data['current_password']
        new_password = data['new_password']
        user_id = session.get('user_id')
        username = session.get('username')
        

        

        
        # 验证新密码长度
        if len(new_password) < 6:
            return jsonify({'success': False, 'message': '新密码长度至少6位'})
        
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor(dictionary=True)
            
            # 获取当前用户信息
            if user_id:
                cursor.execute("SELECT password FROM employees WHERE id = %s", (user_id,))
            else:
                cursor.execute("SELECT password FROM employees WHERE username = %s", (username,))
            user = cursor.fetchone()
            
            if not user:
                return jsonify({'success': False, 'message': '用户不存在'})
            
            # 验证当前密码
            if not check_password_hash(user['password'], current_password):
                return jsonify({'success': False, 'message': '当前密码错误'})
            
            # 更新密码
            new_password_hash = generate_password_hash(new_password)
            if user_id:
                cursor.execute("UPDATE employees SET password = %s WHERE id = %s", 
                             (new_password_hash, user_id))
            else:
                cursor.execute("UPDATE employees SET password = %s WHERE username = %s", 
                             (new_password_hash, username))
            
            connection.commit()
            
            return jsonify({'success': True, 'message': '密码修改成功'})
        else:
            return jsonify({'success': False, 'message': '数据库连接失败'})
            
    except Error as e:
        return jsonify({'success': False, 'message': f'修改失败: {str(e)}'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'系统错误: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 获取所有账单信息
def get_all_bill_info(page=1, per_page=15):
    connection = get_db_connection()
    if connection:
        try:
            cursor = connection.cursor(dictionary=True)
            
            # 计算偏移量
            offset = (page - 1) * per_page
            
            # 获取总记录数
            cursor.execute("SELECT COUNT(*) as total FROM bill_info")
            total_count = cursor.fetchone()['total']
            
            # 获取分页数据
            cursor.execute("SELECT * FROM bill_info ORDER BY created_at DESC LIMIT %s OFFSET %s", 
                         (per_page, offset))
            bill_info = cursor.fetchall()
            
            # 计算总页数
            total_pages = (total_count + per_page - 1) // per_page
            
            return {
                'data': bill_info,
                'total': total_count,
                'page': page,
                'per_page': per_page,
                'total_pages': total_pages
            }
        except Error as e:
            print(f"获取账单信息列表错误: {e}")
            return {'data': [], 'total': 0, 'page': 1, 'per_page': per_page, 'total_pages': 0}
        finally:
            cursor.close()
            connection.close()
    return {'data': [], 'total': 0, 'page': 1, 'per_page': per_page, 'total_pages': 0}

@app.route('/bill_info')
def bill_info():
    """账单信息管理页面"""
    if 'loggedin' not in session:
        return redirect(url_for('login'))
    
    # 检查权限
    user_id = session.get('user_id')
    permissions = get_user_permissions(user_id)
    
    if 'bill_management' not in permissions:
        flash('您没有访问账单信息管理的权限！', 'error')
        return redirect(url_for('dashboard'))
    
    # 获取第一页数据用于初始显示
    bill_info_result = get_all_bill_info(1, 15)
    return render_template('bill_info.html', 
                         username=session.get('user_name', session['username']), 
                         permissions=permissions,
                         bill_info=bill_info_result['data'],
                         pagination={
                             'total': bill_info_result['total'],
                             'page': bill_info_result['page'],
                             'per_page': bill_info_result['per_page'],
                             'total_pages': bill_info_result['total_pages']
                         })

@app.route('/get_bill_info')
def get_bill_info():
    """获取账单信息列表API"""
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        # 获取分页参数
        page = request.args.get('page', 1, type=int)
        per_page = request.args.get('per_page', 15, type=int)
        
        # 限制每页最大数量
        if per_page > 100:
            per_page = 100
        
        bill_info_result = get_all_bill_info(page, per_page)
        return jsonify({'success': True, **bill_info_result})
    except Exception as e:
        return jsonify({'success': False, 'message': f'获取数据失败: {str(e)}'})

@app.route('/get_bill_info/<int:bill_id>')
def get_single_bill_info(bill_id):
    """获取单个账单信息API"""
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    connection = get_db_connection()
    if not connection:
        return jsonify({'success': False, 'message': '数据库连接失败'})
    
    try:
        cursor = connection.cursor(dictionary=True)
        cursor.execute("SELECT * FROM bill_info WHERE id = %s", (bill_id,))
        bill = cursor.fetchone()
        
        if bill:
            return jsonify({'success': True, 'data': bill})
        else:
            return jsonify({'success': False, 'message': '未找到账单信息'})
            
    except Error as e:
        return jsonify({'success': False, 'message': f'获取数据失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

@app.route('/add_bill_info', methods=['POST'])
def add_bill_info():
    """添加账单信息"""
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    connection = get_db_connection()
    if not connection:
        return jsonify({'success': False, 'message': '数据库连接失败'})
    
    cursor = None
    try:
        data = request.get_json() or request.form.to_dict()
        print(f"Received data: {data}")  # 调试日志
        
        # 验证数据
        is_valid, errors = validate_bill_data(data, connection)
        print(f"Validation result: {is_valid}, errors: {errors}")  # 调试日志
        if not is_valid:
            return jsonify({'success': False, 'message': '; '.join(errors)})
        
        cursor = connection.cursor()
        
        # 插入新账单信息
        insert_query = """
        INSERT INTO bill_info (mail_class, des, route_info, flight_no, quote, carry_code)
        VALUES (%s, %s, %s, %s, %s, %s)
        """
        
        # 安全处理quote字段
        quote_value = data.get('quote')
        if quote_value:
            try:
                quote_float = float(quote_value)
            except (ValueError, TypeError):
                quote_float = None
        else:
            quote_float = None
            
        bill_data = (
            data.get('mail_class', ''),
            data.get('des', ''),
            data.get('route_info', ''),
            data.get('flight_no', ''),
            quote_float,
            data.get('carry_code', '')
        )
        
        cursor.execute(insert_query, bill_data)
        connection.commit()
        
        return jsonify({'success': True, 'message': '账单信息添加成功'})
        
    except Error as e:
        return jsonify({'success': False, 'message': f'添加失败: {str(e)}'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'系统错误: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

@app.route('/update_bill_info/<int:bill_id>', methods=['PUT'])
def update_bill_info(bill_id):
    """更新账单信息"""
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    connection = get_db_connection()
    if not connection:
        return jsonify({'success': False, 'message': '数据库连接失败'})
    
    try:
        data = request.get_json() or request.form.to_dict()
        
        # 验证数据（排除当前记录ID）
        is_valid, errors = validate_bill_data(data, connection, exclude_id=bill_id)
        if not is_valid:
            return jsonify({'success': False, 'message': '; '.join(errors)})
        
        cursor = connection.cursor()
        
        # 更新账单信息
        update_query = """
        UPDATE bill_info 
        SET mail_class = %s, des = %s, route_info = %s, flight_no = %s, quote = %s, carry_code = %s
        WHERE id = %s
        """
        
        # 安全处理quote字段
        quote_value = data.get('quote')
        if quote_value:
            try:
                quote_float = float(quote_value)
            except (ValueError, TypeError):
                quote_float = None
        else:
            quote_float = None
            
        bill_data = (
            data.get('mail_class', ''),
            data.get('des', ''),
            data.get('route_info', ''),
            data.get('flight_no', ''),
            quote_float,
            data.get('carry_code', ''),
            bill_id
        )
        
        cursor.execute(update_query, bill_data)
        connection.commit()
        
        if cursor.rowcount > 0:
            return jsonify({'success': True, 'message': '账单信息更新成功'})
        else:
            return jsonify({'success': False, 'message': '未找到要更新的账单信息'})
            
    except Error as e:
        return jsonify({'success': False, 'message': f'更新失败: {str(e)}'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'系统错误: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

@app.route('/delete_bill_info/<int:bill_id>', methods=['DELETE'])
def delete_bill_info(bill_id):
    """删除账单信息"""
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    connection = get_db_connection()
    if not connection:
        return jsonify({'success': False, 'message': '数据库连接失败'})
    
    try:
        cursor = connection.cursor()
        
        cursor.execute("DELETE FROM bill_info WHERE id = %s", (bill_id,))
        connection.commit()
        
        if cursor.rowcount > 0:
            return jsonify({'success': True, 'message': '账单信息删除成功'})
        else:
            return jsonify({'success': False, 'message': '未找到要删除的账单信息'})
            
    except Error as e:
        return jsonify({'success': False, 'message': f'删除失败: {str(e)}'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'系统错误: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()



@app.route('/import_bill_excel', methods=['POST'])
def import_bill_excel():
    """导入Excel账单信息"""
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': '未选择文件'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'message': '未选择文件'})
    
    if not allowed_file(file.filename):
        return jsonify({'success': False, 'message': '文件格式不支持，请上传.xlsx或.xls文件'})
    
    connection = None
    try:
        # 读取Excel文件
        df = pd.read_excel(file)
        
        # 验证表头
        required_columns = ['邮件类型', '目的地', '路由信息', '航班号', '报价', '承运代码']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            return jsonify({
                'success': False, 
                'message': f'Excel表头缺少必需字段: {", ".join(missing_columns)}'
            })
        
        # 数据预处理
        df = df[required_columns]  # 只保留需要的列
        df = df.dropna(how='all')  # 删除完全空白的行
        
        if df.empty:
            return jsonify({'success': False, 'message': 'Excel文件中没有有效数据'})
        
        # 获取数据库连接
        connection = get_db_connection()
        if not connection:
            return jsonify({'success': False, 'message': '数据库连接失败'})
        
        # 开始事务
        connection.start_transaction()
        cursor = connection.cursor()
        
        # 批量验证所有数据
        validation_errors = []
        processed_data = []
        existing_combinations = set()  # 用于检查Excel内部重复
        
        for index, row in df.iterrows():
            row_num = index + 2  # Excel行号（从第2行开始，第1行是表头）
            
            # 构造数据字典
            data = {
                'mail_class': str(row['邮件类型']).strip() if pd.notna(row['邮件类型']) else '',
                'des': str(row['目的地']).strip().upper() if pd.notna(row['目的地']) else '',
                'route_info': str(row['路由信息']) if pd.notna(row['路由信息']) else '',
                'flight_no': str(row['航班号']) if pd.notna(row['航班号']) else '',
                'quote': row['报价'] if pd.notna(row['报价']) else None,
                'carry_code': str(row['承运代码']) if pd.notna(row['承运代码']) else ''
            }
            
            # 验证单行数据
            is_valid, errors = validate_bill_data(data, connection)
            if not is_valid:
                for error in errors:
                    validation_errors.append(f'第{row_num}行: {error}')
                continue
            
            # 检查Excel内部重复
            combination_key = (data['mail_class'], data['des'])
            if combination_key in existing_combinations:
                validation_errors.append(f'第{row_num}行: 该目的地已有报价')
                continue
            existing_combinations.add(combination_key)
            
            processed_data.append(data)
        
        # 如果有验证错误，回滚事务并返回错误
        if validation_errors:
            connection.rollback()
            return jsonify({
                'success': False, 
                'message': '数据验证失败，导入已取消:\n' + '\n'.join(validation_errors[:10])  # 最多显示10个错误
            })
        
        # 所有数据验证通过，批量插入
        insert_query = """
        INSERT INTO bill_info (mail_class, des, route_info, flight_no, quote, carry_code)
        VALUES (%s, %s, %s, %s, %s, %s)
        """
        
        for data in processed_data:
            bill_data = (
                data['mail_class'],
                data['des'],
                data['route_info'],
                data['flight_no'],
                float(data['quote']) if data['quote'] is not None else None,
                data['carry_code']
            )
            cursor.execute(insert_query, bill_data)
        
        # 提交事务
        connection.commit()
        
        return jsonify({
            'success': True, 
            'message': f'导入成功，共导入{len(processed_data)}条记录',
            'imported_count': len(processed_data)
        })
        
    except Exception as e:
        if connection:
            connection.rollback()
        return jsonify({'success': False, 'message': f'导入失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if connection:
            connection.close()

# 获取邮件数据列表API
@app.route('/api/mail_data', methods=['GET'])
def get_mail_data():
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        page = int(request.args.get('page', 1))
        per_page = int(request.args.get('per_page', 15))
        
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor(dictionary=True)
            
            # 计算偏移量
            offset = (page - 1) * per_page
            
            # 获取总记录数
            cursor.execute("SELECT COUNT(*) as total FROM mail_data")
            total_count = cursor.fetchone()['total']
            
            # 获取分页数据
            cursor.execute("""
                SELECT mail_id, mail_receptacleNo, mail_originPost, mail_destPost, 
                       mail_dest, mail_recTime, mail_upliftTime, mail_arriveTime, 
                       mail_deliverTime, mail_routeInfo, mail_flightInfo, mail_weight, 
                       mail_quote, mail_charge, mail_carrCode, created_at
                FROM mail_data 
                ORDER BY created_at DESC 
                LIMIT %s OFFSET %s
            """, (per_page, offset))
            mail_data = cursor.fetchall()
            
            # 计算总页数
            total_pages = (total_count + per_page - 1) // per_page
            
            return jsonify({
                'success': True,
                'data': mail_data,
                'total': total_count,
                'page': page,
                'per_page': per_page,
                'total_pages': total_pages
            })
        else:
            return jsonify({'success': False, 'message': '数据库连接失败'})
            
    except Exception as e:
        return jsonify({'success': False, 'message': f'获取数据失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if 'connection' in locals() and connection:
            connection.close()

# 添加邮件数据API
# 验证邮件类型格式（与账单信息管理相同）
def validate_mail_class_for_mail_data(mail_class):
    if not mail_class:
        return False, '邮件类型不能为空'
    if mail_class not in ['TY', 'PY']:
        return False, '邮件类型字段不符合要求'
    return True, ''

# 验证总包号格式
def validate_receptacle_no(receptacle_no):
    if not receptacle_no:
        return False, '总包号不能为空'
    
    if len(receptacle_no) != 29:
        return False, '总包号长度必须为29位'
    
    letter_part = receptacle_no[:15]
    number_part = receptacle_no[15:]
    
    # 检查前15位是否为字母
    if not letter_part.isalpha():
        return False, '总包号前15位必须为字母'
    
    # 检查后14位是否为数字
    if not number_part.isdigit():
        return False, '总包号后14位必须为数字'
    
    return True, ''

# 验证到达地是否存在于bill_info表的des字段
@app.route('/api/validate_destination', methods=['POST'])
def validate_destination():
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        data = request.get_json()
        destination = data.get('destination', '').strip()
        
        if not destination:
            return jsonify({'success': False, 'message': '到达地不能为空'})
        
        connection = get_db_connection()
        cursor = connection.cursor()
        
        # 查询bill_info表中是否存在该到达地
        cursor.execute("SELECT COUNT(*) FROM bill_info WHERE des = %s", (destination,))
        count = cursor.fetchone()[0]
        
        if count > 0:
            return jsonify({'success': True, 'message': '到达地验证通过'})
        else:
            return jsonify({'success': False, 'message': '不存在该到达地'})
            
    except Exception as e:
        return jsonify({'success': False, 'message': f'验证失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if 'connection' in locals() and connection:
            connection.close()

@app.route('/api/mail_data', methods=['POST'])
def add_mail_data():
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        data = request.get_json()
        
        # 验证邮件类型格式
        mail_class = data.get('Mail_class', '').strip().upper()
        is_valid, error_msg = validate_mail_class_for_mail_data(mail_class)
        if not is_valid:
            return jsonify({'success': False, 'message': error_msg})
        
        # 将邮件类型转换为大写
        data['Mail_class'] = mail_class
        
        # 验证总包号格式
        receptacle_no = data.get('mail_receptacleNo', '')
        is_valid, error_msg = validate_receptacle_no(receptacle_no)
        if not is_valid:
            return jsonify({'success': False, 'message': error_msg})
        
        # 将总包号前15位转换为大写
        data['mail_receptacleNo'] = receptacle_no[:15].upper() + receptacle_no[15:]
        
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            
            insert_query = """
                INSERT INTO mail_data (Mail_class, mail_receptacleNo, mail_originPost, mail_destPost, 
                                     mail_dest, mail_recTime, mail_upliftTime, mail_arriveTime, 
                                     mail_deliverTime, mail_routeInfo, mail_flightInfo, 
                                     mail_weight, mail_quote, mail_charge, mail_carrCode)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """
            
            cursor.execute(insert_query, (
                data.get('Mail_class'),
                data.get('mail_receptacleNo'),
                data.get('mail_originPost'),
                data.get('mail_destPost'),
                data.get('mail_dest'),
                data.get('mail_recTime'),
                data.get('mail_upliftTime'),
                data.get('mail_arriveTime'),
                data.get('mail_deliverTime'),
                data.get('mail_routeInfo'),
                data.get('mail_flightInfo'),
                data.get('mail_weight'),
                data.get('mail_quote'),
                data.get('mail_charge'),
                data.get('mail_carrCode')
            ))
            
            connection.commit()
            return jsonify({'success': True, 'message': '邮件数据添加成功'})
        else:
            return jsonify({'success': False, 'message': '数据库连接失败'})
            
    except Exception as e:
        return jsonify({'success': False, 'message': f'添加失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if 'connection' in locals() and connection:
            connection.close()

# 更新邮件数据API
@app.route('/api/mail_data/<int:mail_id>', methods=['PUT'])
def update_mail_data(mail_id):
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        data = request.get_json()
        
        # 验证邮件类型格式
        mail_class = data.get('Mail_class', '').strip().upper()
        is_valid, error_msg = validate_mail_class_for_mail_data(mail_class)
        if not is_valid:
            return jsonify({'success': False, 'message': error_msg})
        
        # 将邮件类型转换为大写
        data['Mail_class'] = mail_class
        
        # 验证总包号格式
        receptacle_no = data.get('mail_receptacleNo', '')
        is_valid, error_msg = validate_receptacle_no(receptacle_no)
        if not is_valid:
            return jsonify({'success': False, 'message': error_msg})
        
        # 将总包号前15位转换为大写
        data['mail_receptacleNo'] = receptacle_no[:15].upper() + receptacle_no[15:]
        
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            
            update_query = """
                UPDATE mail_data SET 
                    Mail_class = %s, mail_receptacleNo = %s, mail_originPost = %s, mail_destPost = %s,
                    mail_dest = %s, mail_recTime = %s, mail_upliftTime = %s,
                    mail_arriveTime = %s, mail_deliverTime = %s, mail_routeInfo = %s,
                    mail_flightInfo = %s, mail_weight = %s, mail_quote = %s,
                    mail_charge = %s, mail_carrCode = %s
                WHERE mail_id = %s
            """
            
            cursor.execute(update_query, (
                data.get('Mail_class'),
                data.get('mail_receptacleNo'),
                data.get('mail_originPost'),
                data.get('mail_destPost'),
                data.get('mail_dest'),
                data.get('mail_recTime'),
                data.get('mail_upliftTime'),
                data.get('mail_arriveTime'),
                data.get('mail_deliverTime'),
                data.get('mail_routeInfo'),
                data.get('mail_flightInfo'),
                data.get('mail_weight'),
                data.get('mail_quote'),
                data.get('mail_charge'),
                data.get('mail_carrCode'),
                mail_id
            ))
            
            connection.commit()
            return jsonify({'success': True, 'message': '邮件数据更新成功'})
        else:
            return jsonify({'success': False, 'message': '数据库连接失败'})
            
    except Exception as e:
        return jsonify({'success': False, 'message': f'更新失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if 'connection' in locals() and connection:
            connection.close()

# 删除邮件数据API
@app.route('/api/mail_data/<int:mail_id>', methods=['DELETE'])
def delete_mail_data(mail_id):
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        connection = get_db_connection()
        if connection:
            cursor = connection.cursor()
            
            cursor.execute("DELETE FROM mail_data WHERE mail_id = %s", (mail_id,))
            connection.commit()
            
            if cursor.rowcount > 0:
                return jsonify({'success': True, 'message': '邮件数据删除成功'})
            else:
                return jsonify({'success': False, 'message': '未找到要删除的数据'})
        else:
            return jsonify({'success': False, 'message': '数据库连接失败'})
            
    except Exception as e:
        return jsonify({'success': False, 'message': f'删除失败: {str(e)}'})
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if 'connection' in locals() and connection:
            connection.close()

@app.route('/logout')
def logout():
    """退出登录"""
    session.pop('loggedin', None)
    session.pop('username', None)
    flash('已成功退出登录！', 'info')
    return redirect(url_for('login'))

# 初始化数据库表结构
def init_bill_info_table():
    """初始化账单信息表"""
    connection = get_db_connection()
    if connection:
        try:
            cursor = connection.cursor()
            
            # 创建账单信息表
            create_bill_info_table = """
            CREATE TABLE IF NOT EXISTS bill_info (
                id INT AUTO_INCREMENT PRIMARY KEY,
                mail_class VARCHAR(100) NOT NULL COMMENT '邮件类型',
                des VARCHAR(200) NOT NULL COMMENT '目的地',
                route_info VARCHAR(500) COMMENT '路由信息',
                flight_no VARCHAR(100) COMMENT '航班信息',
                quote DECIMAL(10,2) COMMENT '报价',
                carry_code VARCHAR(100) COMMENT '运能编码',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """
            
            cursor.execute(create_bill_info_table)
            
            # 创建角色表
            create_roles_table = """
            CREATE TABLE IF NOT EXISTS roles (
                id INT AUTO_INCREMENT PRIMARY KEY,
                role_name VARCHAR(50) UNIQUE NOT NULL,
                permissions TEXT COMMENT '权限列表，逗号分隔',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """
            
            cursor.execute(create_roles_table)
            
            # 检查是否已存在默认角色
            cursor.execute("SELECT COUNT(*) FROM roles WHERE role_name = '管理员'")
            admin_role_exists = cursor.fetchone()[0]
            
            # 创建产品信息表
            create_products_table = """
            CREATE TABLE IF NOT EXISTS products (
                product_id INT AUTO_INCREMENT PRIMARY KEY,
                product_code VARCHAR(50) UNIQUE NOT NULL COMMENT '产品代码',
                product_name VARCHAR(200) NOT NULL COMMENT '产品名称',
                product_type VARCHAR(50) NOT NULL COMMENT '产品类型',
                unit_price DECIMAL(10,2) NOT NULL COMMENT '单价',
                status VARCHAR(20) DEFAULT '启用' COMMENT '状态',
                description TEXT COMMENT '产品描述',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """
            
            cursor.execute(create_products_table)
            
            # 创建邮件数据表
            create_mail_data_table = """
            CREATE TABLE IF NOT EXISTS mail_data (
                mail_id INT AUTO_INCREMENT PRIMARY KEY,
                Mail_class VARCHAR(100) NOT NULL COMMENT '邮件类型',
                mail_receptacleNo VARCHAR(100) NOT NULL COMMENT '总包号',
                mail_originPost VARCHAR(200) NOT NULL COMMENT '始发局',
                mail_destPost VARCHAR(200) NOT NULL COMMENT '寄达局',
                mail_dest VARCHAR(200) NOT NULL COMMENT '到达地',
                mail_recTime DATETIME COMMENT '接收时间',
                mail_upliftTime DATETIME COMMENT '启运时间',
                mail_arriveTime DATETIME COMMENT '到达时间',
                mail_deliverTime DATETIME COMMENT '交邮时间',
                mail_routeInfo VARCHAR(500) COMMENT '收费路由',
                mail_flightInfo VARCHAR(200) COMMENT '航班',
                mail_weight DECIMAL(10,3) COMMENT '重量',
                mail_quote DECIMAL(10,2) COMMENT '费率',
                mail_charge DECIMAL(10,2) COMMENT '金额',
                mail_carrCode VARCHAR(100) COMMENT '运能编码',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """
            
            cursor.execute(create_mail_data_table)
            
            if admin_role_exists == 0:
                # 插入默认角色
                insert_roles = """
                INSERT INTO roles (role_name, permissions) VALUES 
                ('管理员', 'dashboard,employees,roles,products,bill_management,mail_data'),
                ('普通用户', 'bill_management')
                """
                cursor.execute(insert_roles)
            
            # 修改员工表，添加role_id字段
            try:
                cursor.execute("ALTER TABLE employees ADD COLUMN role_id INT DEFAULT 2")
                cursor.execute("ALTER TABLE employees ADD FOREIGN KEY (role_id) REFERENCES roles(id)")
            except Error as e:
                if "Duplicate column name" not in str(e):
                    print(f"修改员工表结构时出错: {e}")
            
            # 更新现有管理员用户的角色
            cursor.execute("UPDATE employees SET role_id = 1 WHERE role_type = '管理员'")
            cursor.execute("UPDATE employees SET role_id = 2 WHERE role_type = '一般员工'")
            
            connection.commit()
            print("账单信息表和角色表初始化成功")
            
        except Error as e:
            print(f"初始化账单信息表错误: {e}")
        finally:
            cursor.close()
            connection.close()

if __name__ == '__main__':
    init_database()
    init_bill_info_table()
    app.run(debug=True, host='0.0.0.0', port=5000)