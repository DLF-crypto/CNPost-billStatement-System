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
        if connection:
            cursor.close()
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
        if connection:
            cursor.close()
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
        if connection:
            cursor.close()
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
        if connection:
            cursor.close()
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
        if connection:
            cursor.close()
            connection.close()

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
        if connection:
            cursor.close()
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
        if connection:
            cursor.close()
            connection.close()

# 获取所有账单信息
def get_all_bill_info():
    connection = get_db_connection()
    if connection:
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute("SELECT * FROM bill_info ORDER BY created_at DESC")
            bill_info = cursor.fetchall()
            return bill_info
        except Error as e:
            print(f"获取账单信息列表错误: {e}")
            return []
        finally:
            cursor.close()
            connection.close()
    return []

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
    
    bill_info_list = get_all_bill_info()
    return render_template('bill_info.html', 
                         username=session['username'], 
                         permissions=permissions,
                         bill_info=bill_info_list)

@app.route('/add_bill_info', methods=['POST'])
def add_bill_info():
    """添加账单信息"""
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    connection = get_db_connection()
    if not connection:
        return jsonify({'success': False, 'message': '数据库连接失败'})
    
    try:
        data = request.get_json()
        cursor = connection.cursor()
        
        # 插入新账单信息
        insert_query = """
        INSERT INTO bill_info (mail_class, des, route_info, flight_no, quote, carry_code)
        VALUES (%s, %s, %s, %s, %s, %s)
        """
        
        bill_data = (
            data['mail_class'],
            data['des'],
            data.get('route_info', ''),
            data.get('flight_no', ''),
            data.get('quote', None),
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
        if connection:
            cursor.close()
            connection.close()

@app.route('/update_bill_info', methods=['POST'])
def update_bill_info():
    """更新账单信息"""
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    connection = get_db_connection()
    if not connection:
        return jsonify({'success': False, 'message': '数据库连接失败'})
    
    try:
        data = request.get_json()
        cursor = connection.cursor()
        bill_id = data['id']
        
        # 更新账单信息
        update_query = """
        UPDATE bill_info 
        SET mail_class = %s, des = %s, route_info = %s, flight_no = %s, quote = %s, carry_code = %s
        WHERE id = %s
        """
        
        bill_data = (
            data['mail_class'],
            data['des'],
            data.get('route_info', ''),
            data.get('flight_no', ''),
            data.get('quote', None),
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
        if connection:
            cursor.close()
            connection.close()

@app.route('/delete_bill_info', methods=['POST'])
def delete_bill_info():
    """删除账单信息"""
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    connection = get_db_connection()
    if not connection:
        return jsonify({'success': False, 'message': '数据库连接失败'})
    
    try:
        data = request.get_json()
        cursor = connection.cursor()
        bill_id = data['id']
        
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
        if connection:
            cursor.close()
            connection.close()

@app.route('/import_bill_excel', methods=['POST'])
def import_bill_excel():
    """导入Excel账单信息"""
    if 'loggedin' not in session:
        return jsonify({'success': False, 'message': '请先登录'})
    
    try:
        # 检查是否有文件上传
        if 'excel_file' not in request.files:
            return jsonify({'success': False, 'message': '请选择要上传的Excel文件'})
        
        file = request.files['excel_file']
        
        # 检查文件名
        if file.filename == '':
            return jsonify({'success': False, 'message': '请选择要上传的Excel文件'})
        
        if file and allowed_file(file.filename):
            # 保存文件
            filename = secure_filename(file.filename)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"{timestamp}_{filename}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            try:
                # 读取Excel文件
                df = pd.read_excel(filepath)
                
                # 检查必需的列
                required_columns = ['邮件类型', '目的地']
                optional_columns = ['路由信息', '航班信息', '报价', '运能编码']
                
                missing_columns = [col for col in required_columns if col not in df.columns]
                if missing_columns:
                    os.remove(filepath)  # 删除上传的文件
                    return jsonify({'success': False, 'message': f'Excel文件缺少必需的列: {", ".join(missing_columns)}'})
                
                # 连接数据库
                connection = get_db_connection()
                if not connection:
                    os.remove(filepath)
                    return jsonify({'success': False, 'message': '数据库连接失败'})
                
                cursor = connection.cursor()
                
                success_count = 0
                error_count = 0
                error_messages = []
                
                # 逐行处理数据
                for index, row in df.iterrows():
                    try:
                        # 获取必需字段
                        mail_class = str(row['邮件类型']).strip() if pd.notna(row['邮件类型']) else ''
                        des = str(row['目的地']).strip() if pd.notna(row['目的地']) else ''
                        
                        if not mail_class or not des:
                            error_count += 1
                            error_messages.append(f'第{index+2}行: 邮件类型和目的地不能为空')
                            continue
                        
                        # 获取可选字段
                        route_info = str(row['路由信息']).strip() if '路由信息' in df.columns and pd.notna(row['路由信息']) else None
                        flight_no = str(row['航班信息']).strip() if '航班信息' in df.columns and pd.notna(row['航班信息']) else None
                        carry_code = str(row['运能编码']).strip() if '运能编码' in df.columns and pd.notna(row['运能编码']) else None
                        
                        # 处理报价字段
                        quote = None
                        if '报价' in df.columns and pd.notna(row['报价']):
                            try:
                                quote = float(row['报价'])
                            except (ValueError, TypeError):
                                quote = None
                        
                        # 插入数据库
                        insert_query = """
                            INSERT INTO bill_info (mail_class, des, route_info, flight_no, quote, carry_code)
                            VALUES (%s, %s, %s, %s, %s, %s)
                        """
                        cursor.execute(insert_query, (mail_class, des, route_info, flight_no, quote, carry_code))
                        success_count += 1
                        
                    except Exception as row_error:
                        error_count += 1
                        error_messages.append(f'第{index+2}行: {str(row_error)}')
                        continue
                
                connection.commit()
                
                # 删除上传的文件
                os.remove(filepath)
                
                # 返回结果
                if success_count > 0:
                    message = f'导入完成！成功导入 {success_count} 条记录'
                    if error_count > 0:
                        message += f'，失败 {error_count} 条记录'
                        if len(error_messages) <= 5:  # 只显示前5个错误
                            message += f'\n错误详情：\n' + '\n'.join(error_messages[:5])
                    return jsonify({'success': True, 'message': message})
                else:
                    message = '导入失败，没有成功导入任何记录'
                    if error_messages:
                        message += f'\n错误详情：\n' + '\n'.join(error_messages[:5])
                    return jsonify({'success': False, 'message': message})
                    
            except Exception as excel_error:
                # 删除上传的文件
                if os.path.exists(filepath):
                    os.remove(filepath)
                return jsonify({'success': False, 'message': f'Excel文件处理错误: {str(excel_error)}'})
            finally:
                if connection:
                    cursor.close()
                    connection.close()
                
        else:
            return jsonify({'success': False, 'message': '请上传有效的Excel文件(.xlsx或.xls格式)'})
            
    except Exception as e:
        return jsonify({'success': False, 'message': f'导入失败: {str(e)}'})

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
                route_info TEXT COMMENT '路由信息',
                flight_no VARCHAR(50) COMMENT '航班信息',
                quote DECIMAL(10,2) COMMENT '报价',
                carry_code VARCHAR(50) COMMENT '运能编码',
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
            
            if admin_role_exists == 0:
                # 插入默认角色
                insert_roles = """
                INSERT INTO roles (role_name, permissions) VALUES 
                ('管理员', 'employees,roles,bill_management'),
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