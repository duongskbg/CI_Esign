# -*- coding: utf-8 -*-
"""
Created on Mon May 23 17:23:24 2022

@author: V1008647
"""
import os,json,shutil,sqlite3,re,asyncio,requests,random,logging,http.client
from flask import Flask, request, redirect, url_for, render_template, session, jsonify, send_file, abort
from authlib.integrations.flask_client import OAuth
from auth_decorator import login_required
from werkzeug.http import HTTP_STATUS_CODES
from flask_session import Session
from datetime import datetime, timedelta
from functools import wraps
import pandas as pd
http.client.HTTPConnection.debuglevel = 1
logging.basicConfig()
logging.getLogger().setLevel(logging.DEBUG)
requests_log = logging.getLogger("requests.packages.urllib3")
requests_log.setLevel(logging.DEBUG)
requests_log.propagate = True

pd.set_option('display.width', 1000)
pd.set_option('colheader_justify', 'center')

def async_action(f):
    @wraps(f)
    def wrapped(*args, **kwargs):
        return asyncio.run(f(*args, **kwargs))
    return wrapped

app = Flask(__name__)
# Session config
app.config['SESSION_COOKIE_NAME'] = 'login-session'
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes=45)
app.config["SESSION_PERMANENT"] = False
app.config["SESSION_TYPE"] = "filesystem"
Session(app)

# oAuth Setup
oauth = OAuth(app)
oauthV2 = oauth.register(
    name='VN AI System',
    client_assertion_type='OAuth2.Client.FoxconnClient',
    client_id="deb912cf17ded68caa338911982d994c",
    client_secret="lds2qpi76qt7svsr8nlcy5pjg17qv80qked3wqu5e5a0ygmqfcoqhx8iwloe35pa",
    # This is only needed if using openId to fetch user info
    userinfo_endpoint='https://lh-account.cnsbg.efoxconn.com/openidconnect/v1/userinfo',
    api_base_url='http://10.220.40.75:5000/authorize',
    authorize_url='https://lh-account.cnsbg.efoxconn.com/oauth2/v3/auth',  # Sever Long Hoa
    access_token_url='https://lh-account.cnsbg.efoxconn.com/oauth2/v3/token',  # Sever Long Hoa
    access_token_params=None,
    authorize_params=None,
    client_kwargs={'scope': 'profile'},
)

black_list = ['10.220.35.3']
@app.before_request
def before_request():
    session.permanent = True
    # phiên làm việc sẽ tự động xóa sau 45p nếu không có thêm bất cứ request nào lên server.
    app.permanent_session_lifetime = timedelta(minutes=45)
    ip = request.environ.get('REMOTE_ADDR')
    if ip in black_list:
        abort(403)

@app.errorhandler(404)
def not_found(error):
    return render_template('404.html'), 404

# ========================API=================================
@app.route('/api/get_language/<language>', methods = ['GET','POST'])
def get_language(language):
    data = []
    if language == 'vn':
        data = {'text1': 'Danh sách', 'text2': 'Tiêu đề', 'text3': 'Trạng Thái','text4': 'Người Tạo','text5': 'Thời Gian Tạo','text6': 'Trang trước', 'text7': 'Trang sau','home':'Trang chủ','create': 'Tạo mới','no_sign': 'Ký Thường','ex_sign':'ký qua file excel','sign_app':'Ký đơn','sign_no':'Đơn Thường','sign_ex':'Đơn ký bằng Excel','my_app':'Đơn của tôi','all':'Tất cả đơn','waiting':'Đơn chờ ký','sign':'Đơn đã được ký','reject':'Đơn bị trả về','sign_app':'Ký đơn','sign_no':'Đơn Thường','sign_ex':'Đơn ký bằng Excel','so_sign':'Ký đơn', 'manv':'Mã nhân viên','content':'Nội Dung','ch_sign':'Thêm Người Ký','btn_pre':'Xem Lại','ch_file':'Chọn File','small_t':'Đây là những file cần thiết để so sánh trong quá trình ký.'\
                ,'update_time':'Cập nhật','title_edit':'Chỉnh sửa','reject_time':'Từ chối','approve_time':'Thông qua','title_alert':'Thông báo','title_signer':'Người ký','button_id1':'Quay lại','button_id':'Gửi đi','order_id':'Mã đơn','full_name':'Tên đầy đủ','order_title':'Tiêu đề đơn','order_content':'Nội dung đơn','attached_file':'Tệp đính kèm','create_at':'Tạo lúc: ','reason':'Lý do:','lb_app':'Thông qua','lb_rej':'Từ chối','button_id3':'Đổi người ký','button_id2':'Từ Chối','button_id4':'Ký','button_id5':'Đổi','note':'Ghi chú','ck_all':'Chọn tất cả','atten':' Lưu ý:Chỉ có trợ lý của bộ phận mới có quyền làm đơn này','type1':'Loại hình ký','type2':'Mã thẻ người ký','type3':'Người phê duyệt','type4':'Người xin đơn','type5':'Chủ quản cấp phòng','type6':'Chủ quản cấp bộ phận','type7':'Cost xét duyệt','type8':'Chủ quản Cost phê duyệt','co_s':'Người đồng ký'\
                    ,'add':'+ Thêm','submit':'Gửi đơn','search_info':'Tìm kiếm thông tin','user_info':'Thông tin người dùng','btn_return':'Trở về','download':'Tải về','btn_app':'Thông qua'}
    elif language == 'en':
        data = {'text1': 'List Order', 'text2': 'Title', 'text3': 'Status','text4': 'Create By','text5': 'Created Time','text6': 'Previous', 'text7': 'Next','home':'Home','create': 'Create Application','no_sign': 'Normal Sign','ex_sign':'Excel Sign','sign_app':'Sign application','sign_no':'Sign Normal','sign_ex':'Sign Excel','so_sign':'Sign','manv':'Employee','content':'Content','ch_sign':'Choose Signer','btn_pre':'Preview','ch_file':'Choose file','small_t':'These are the files needed for comparison in the signing process.'\
                ,'update_time':'Update Time','title_edit':'Edit','reject_time':'Rejected Time','approve_time':'Approved Time','title_alert':'Alert','title_signer':'Signer','button_id1':'Back','button_id':'Send','order_id':'Order ID','full_name':'Full Name','order_title':'Order Title','my_app':'My Application','order_content':'Order Content','attached_file':'Attached Files','create_at':'Create At:','reason':'Reason:','lb_app':'Approve','lb_rej':'Reject','button_id3':'Change Signer','button_id2': 'Reject','button_id4':'Sign','button_id5':'Change','note': 'Notes','ck_all':'Approve All', 'atten':'Note: Only the assistants of the department have the permission to apply','type1':'Type of signs','type2':'Signer Number','type3':'Signer','type4':'Applicant','type5':'Class supervisor','type6':'Ministerial Director','type7':'Audited by management','type8':'Approved by supervisor','co_s':'co-Signer'\
                    ,'add':'+ Add','submit':'Submit','search_info':'Search Info','user_info':'User information','btn_return':'Return','download':'Download','btn_app':'Approve'}
    else:
        data = {'text1': '清單順序', 'text2': '標題', 'text3': '裝能','text4': '創造者','text5': '創造時間','text6': '上個頁', 'text7': '下個頁','home':'首頁','create': '創造單子','no_sign': '正常簽核','ex_sign':'Excel簽核','sign_app':'單子簽核','sign_no':'正常簽核','sign_ex':'Excel簽核','my_app':'我的單子','all':'所有單子','waiting':'等待簽核','sign':'已簽核','reject':'駁回單子','so_sign':'簽核','manv':'員工','content':'內容','ch_sign':'選擇簽核者','btn_pre':'預覽','ch_file':'選擇文件','small_t':'這些在簽核過程中需要對比的文件。','update_time':'更新時間','title_edit':'維護','reject_time':'駁回時間','approve_time':'批准時間','title_alert':'通知','title_signer':'簽核者','button_id1':'後退','button_id':'發送','order_id':'單子編號','full_name':'全名','order_title':'单標題'\
                ,'order_content':'单内容','attached_file':'附件','create_at':'Create At:','reason':'理由:','lb_app':'批准','lb_rej':'駁回','button_id3':'更改簽核者','button_id2': '駁回','button_id4':'簽核','button_id5':'更改','note': '注意','ck_all':'全部批准', 'atten':'注：部門的助理才有申請的權限。','type1':'簽核類型 ','type2':'簽核工號','type3':'簽核人','type4':'申請人','type5':' 課級主管','type6':'部級主管','type7':'經管審核','type8':'經管主管核准','co_s':'共同签署人','add':'添加','submit':'提交','search_info':'搜索信息','user_info':'用戶信息','btn_return':'返回','download':'下載','btn_app':'批准'}
    session["lang"] = language
    return data

@app.route('/api/create_order', methods=['GET', 'POST'])
def api_create():
    data = 'ss'
    list_signer = []
    data_app = []
    ls = []
    a = ""
    if request.method == 'POST':
        today = datetime.now()
        time = today.strftime("%m/%d/%Y %H:%M:%S")
        data = request.get_json() or {}

        if 'username' not in data or 'list_to_mail' not in data or 'file_type' not in data:
            return bad_request('must include username, email, list_to_mail')
        else:
            username = data['username']
            list_mail = data['list_to_mail']
            order_title = data['order_title']
            descrip = data['description']
            file_name = data['file_name']
            file_type = data['file_type']
            phone = data['phone']
            for i, mail in enumerate(list_mail):
                if mail != "":
                    a += mail + ","
                if (i+1) % 3 == 0:
                    if a != "":
                        ls.append(a)
                        a = ""

            # order_id = str(today.strftime("%Y%m%d%H%M%S")) + str(random.randint(1,999))
            order_id = data['order_id']
            sign_qty = len(ls)
            cr_mail = get_mail(username)
            if get_order('Order_No', 'Create_By', username) != None:
                if order_id in get_order('Order_No', 'Create_By', username):
                    return bad_request('The Order number already exists ')
            for mail in ls:
                if mail.find(cr_mail) != -1:
                    return bad_request('the Sign_email is the same as the creator_email')
                elif re.match(r'\b[A-Za-z0-9._%+-]+@mail.foxconn.com\b', mail) == False or re.match(r'\b[A-Za-z0-9._%+-]+@fii-foxconn.com\b', mail) == False or re.match(r'\b[A-Za-z0-9._%+-]+@foxconn.com\b', mail) == False:
                    return bad_request('the Sign_email is wrong format!' + mail)
            if file_type != "Sign_excel":
                uploads_dir = os.path.join(
                    f"data/{username.upper()}/Normal_Sign", order_id)
            else:
                uploads_dir = os.path.join(
                    f"data/{username.upper()}/Excel_Sign", order_id)
            os.makedirs(uploads_dir, exist_ok=True)
            ex_file = file_name.split(".")[-1]
            if file_name != "":
                taFile = data['taFile']

                if ex_file == "xlsx":
                    pd.read_json(taFile).to_excel(
                        f"{uploads_dir}/{file_name}", index=False, header=False)
                else:
                    pd.read_json(taFile).to_csv(
                        f"{uploads_dir}/{file_name}", index=False, header=False)
            for stt in range(10):
                if stt+1 <= sign_qty:
                    list_signer.append(ls[stt])
                    list_signer.append("")
                    list_signer.append("")
                else:
                    list_signer.append("")
                    list_signer.append("")
                    list_signer.append("")
            # data_ls = str(list_signer)[1:-1]

            data_app.append(order_id)

            data_app.append(username.upper())
            data_app.append(time)
            data_app.append(f'Signed: 0/{sign_qty}')
            data_app.append(sign_qty)
            data_app += list_signer

            data_app.append(username)
            data_app.append(time)
            data_app.append(order_title)
            data_app.append(descrip)
            data_app.append("")
            data_app.append(file_type)
            data_app.append(file_name)
            for i in range(32):
                data_app.append("")
            create_order(str(data_app)[1:-1])
            
            return order_id + " ===== " + send_mail(ls[0], order_title, f'Dear User!<br> You have 1 Order waiting to be signed!<br> Order create by:{username} -- {get_Name_ID(username)}<br>Email: {cr_mail} <br>Phone: {phone} <br>  Please click to this URL to Sign order http://10.220.40.75:5000/esign/{order_id}')

    return jsonify(data)

@app.route('/api/alert', methods=['GET', 'POST'])
def api_alert():
    if request.method == 'POST':
        alert_id = request.form['order_id']
        user_name = request.form['username']
        mail = get_Signer_now(alert_id)
        s_mail = get_mail(user_name)
        ls_title = get_order('Order_Title', 'Order_No', alert_id)[0]
        create_by = get_create_by(alert_id)[0]
        
        send_mail(mail, "(Alert!!!) " + ls_title,
                  f'Dear User!<br>  You have 1 Order waiting to be signed!<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)}  <br> Please click to this URL to Sign order http://10.220.40.75:5000/esign/{alert_id}')
        print(f'{mail}==============================sign')
        send_mail(s_mail, "(Alert!!!) " + ls_title,
                  f'Dear User!<br>  Your Order waiting to be signed! <br> Please Contact {alert_id} -- {mail} to sign the application')
        print("OK")
        return 'done'

@app.route('/api/del/<order_id>', methods=['GET', 'POST'])
def api_del(order_id):
    print(order_id)
    delete_order(order_id)
    print(delete_order(order_id))
    return 'Successfull'

@app.route('/api/sign', methods=['GET', 'POST'])
def api_sign():
    data = 'ss'
    if request.method == 'POST':
        data = request.get_json() or {}
        if 'username' not in data or 'mail' not in data:
            return bad_request('must include username, email')
        else:
            username = data['username']
            mail = data['mail']
            order_id = data['order_id']
            reason = data['reason']
            result = data['result']
            file_type = data['file_type']
            ls_sign = data['ls_sign']
            ls_file = get_order('File_name', 'Order_No', order_id)[0]
            for a in range(10):
                temp = sql_ress(
                    f"SELECT Status_Sign{a+1} FROM Signature_Transaction WHERE Signer_ID{a+1} = '{mail}' AND Order_No = '{order_id}'")
                if temp != None:

                    if temp[0] == "" and (a+1) > 1:
                        temp2 = get_order(
                            f'Status_Sign{a}', 'Order_No', order_id)[0]
                        if temp2[0] != "":
                            lc = a+1
                        else:
                            lc = a

                    elif temp[0] == "" and (a+1) == 1:

                        lc = 1
            if file_type != 'Sign_excel':
                sign_ok(username, mail, order_id, result, lc, reason, '')
            else:
                new_sign = ""
                uploads_dir = os.path.join(
                    f"data/{username.upper()}/Excel_Sign", order_id)
                ls_index = get_order('File_Link', 'Order_No', order_id)[0]
                if ls_index == "":
                    res = f"UPDATE Signature_Transaction SET File_Link = '{ls_sign}' WHERE Order_No = '{order_id}'"
                    update_order(res)
                else:
                    ls_index = ls_index.split(",")
                    for i in ls_sign.split(","):
                        new_sign += str(ls_index[int(i)])+","

                    res = f"UPDATE Signature_Transaction SET File_Link = '{new_sign}' WHERE Order_No = '{order_id}'"
                    update_order(res)
                if not os.path.exists(uploads_dir+"/result.xls"):
                    shutil.copy(uploads_dir+"/"+ls_file.split(",")
                                [0], uploads_dir+"/result.xls")
                wb = pd.read_excel(uploads_dir+"/"+ls_file.split(",")[0])
                ls_sign = ls_sign.split(",")

                if(len(ls_sign) != 0):
                    result = 'approve'
                else:
                    result = 'reject'
                a = wb.iloc[ls_sign]
                da = pd.DataFrame(a)

                da.to_excel(uploads_dir+"/"+ls_file.split(",")[0], index=None)
                sign_ok(username, mail, order_id, result, lc, '', '')
            return order_id + " Sign ok"
    return jsonify(data)

@app.route('/api/get_data/<order_id>')
def get_data(order_id):
    if order_id != 'None':
        if get_order('Signer_Qty', 'Order_No', order_id):
            data = {
                'order_id': order_id
            }
            signer = {}
            qty = get_order('Signer_Qty', 'Order_No', order_id)[0]
            for i in range(qty):
                signer.update({f'Signer{i+1}': {f'sign_id{i+1}': get_order(f'Signer_ID{i+1}', 'Order_No', order_id)[0].split(","), f'Time_sign{i+1}': get_order(
                    f'Time_Signer{i+1}', 'Order_No', order_id)[0], f'Status_Sign{i+1}': get_order(f'Status_Sign{i+1}', 'Order_No', order_id)[0]}})
            data.update(signer)
            data.update({'Order_status': get_order('Order_Status', 'Order_No', order_id)[0], 'Create_Time': get_order(
                'Create_Time', 'Order_No', order_id)[0], 'Reason': get_order('Reason', 'Order_No', order_id)[0]})
            data.update({'Qty': get_order('Signer_Qty', 'Order_No', order_id)[0], 'ls_sign': get_order(
                'File_Link', 'Order_No', order_id)[0], 'sign_time': get_order('Updated_Time', 'Order_No', order_id)[0]})
            return jsonify(data)
        return f'{order_id} không tồn tại!!'
    return 'None'


@app.route('/api/register', methods = ['POST'])
def register_user():
    data = request.get_json()
    print(data)
    username = data['username'] #ten dang nhap
    mail = data['mail'] #email nguoi ky
    order_id = data['order_id'] #ten dau don
    u_name = data['u_name'] #Ten cua nguoi ky
    create_time = data['create_time'] #thoi gian tao don
    phone = data['phone'] #sdt cua nguoi tao don
    name = data['acc_name'] #Ten cua nguoi tao don
    email = data['acc_mail'] #email cua nguoi tao don
    link_url = data['url'] #link vào ký
    send_mail(mail, f"Account Registration",
                  f'Dear {u_name} <br>You have 1 Account waiting to be accepted!!<br> Order No: {order_id}  <br> Account Information:<br>Username: {username}<br> Name: {name}<br>Email: {email}<br>Create Time: {create_time} <br> Phone: {phone}<br>Please click to URL to accept:   {link_url}')
    return 'Done'

@app.route('/api/getlistOrder/<string:mail>')
def st_orget_lider(mail):
    data = {
        'email': mail
    }
    list_order = []
    for a in range(1,10):
        id_order = sql_ress(
            f"SELECT Order_No FROM Signature_Transaction WHERE Signer_ID{a} LIKE '%{mail}%' AND Order_Status NOT LIKE 'True' AND Order_Status NOT LIKE 'Reject'  AND File_Type LIKE 'Sign_excel'")
        if id_order != None:
            for id in id_order:
                temp = get_order(f'Status_Sign{a}', 'Order_No', id)[0]
                if a > 1 and temp == "":
                    temp3 = get_order(f'Status_Sign{a-1}', 'Order_No', id)
                    if temp3[0] != "" and temp3[0] != 'reject':
                        list_order.append(id)
                        print(list_order)
                elif a == 1 and temp == "":
                    list_order.append(id)
    data.update({'order_id': list_order})
    return jsonify(data)

def bad_request(message):
    return error_response(400, message)

def error_response(status_code, message=None):
    payload = {'error': HTTP_STATUS_CODES.get(status_code, 'Unknown error')}
    if message:
        payload['message'] = message
    response = jsonify(payload)
    response.status_code = status_code
    return response

# ===============================END API=====================================

# ===============================Function====================================
def create_connection(db_file):
    conn = None
    try:
        conn = sqlite3.connect(db_file)
    except Exception as e:
        print(e)
    return conn

def sql_res(res):
    connection = sqlite3.connect("db_esign.db")
    cursor = connection.cursor()
    rows = cursor.execute(res).fetchall()
    if rows != []:
        for row in rows:
            return row[0]

def sql_ress(res):
    connection = sqlite3.connect("db_esign.db")
    cursor = connection.cursor()
    rows = cursor.execute(res).fetchall()
    result = []
    if rows != []:
        for row in rows:
            result.append(row[0])
        return result

# function send mail using API
def send_mail(to_mail, title, body):
    url = "https://10.224.81.70:6443/notify-service/api/notify"
    to_mail = to_mail
    title = title
    body = body
    payload = json.dumps({
        "toUser": to_mail,
        "toGroup": "",
        "from": "AI.systems@mail.foxconn.com",
        "message": "{\"title\": \"" + title + "\", \"body\": \"" + body + "\"}",
        "system": "MAIL",
        "source": "WS",
        "type": "TEXT"
    })
    headers = {
        'Content-Type': 'application/json'
    }
    response = requests.request(
        "POST", url, headers=headers, data=payload, verify=False)
    print(str(to_mail) + " send status: " + response.text)
    return response.text


def get_phone(ID):
    api_url = f'http://10.220.40.220:2468/api/get_phone/{ID}'  # URL của API bạn muốn gửi yêu cầu GET
    try:
        response = requests.get(api_url)
        return response.text
    except requests.exceptions.RequestException as e:
        print("Error: ", e)
        return None

def get_Name_ID(username):
    return get_value('Name', 'User_ID', username)

def get_create_by(order_id):
    return get_order('Create_By','Order_No',order_id)

def get_Name_mail(mail):
    return get_value('Name', 'Email', mail)

def get_mail(username):
    return get_value('Email', 'User_ID', username)

def get_ID_mail(mail):
    return get_value('User_ID', 'Email', mail)

def get_value(param, Col_name, Value):
    res = f"SELECT {param} FROM User WHERE {Col_name} like '%{Value}%'"
    if sql_res(res) != "":
        return sql_res(res)
    else:
        return ""

def get_order(param, col_name, value):
    res = f"SELECT {param} FROM Signature_Transaction WHERE {col_name} like '%{value}%'"
    if sql_ress(res) != "":
        return sql_ress(res)
    else:
        return ""

def delete_order(order_id):
    connection = sqlite3.connect("db_esign.db")
    cursor = connection.cursor()
    cursor.execute(
        f"DELETE FROM Signature_Transaction WHERE Order_No = '{order_id}'")
    connection.commit()
    return "done"

def create_order(data):
    connection = sqlite3.connect("db_esign.db")
    cursor = connection.cursor()
    if data != "":
        res = f"INSERT INTO  Signature_Transaction VALUES ({data})"
        cursor.execute(res)
        connection.commit()

def check_user(username, email):
    ck = get_ID_mail(email)
    if username != ck:
        return True
    else:
        return False

def update_order(res):
    connection = sqlite3.connect("db_esign.db")
    cursor = connection.cursor()
    if res != "":
        cursor.execute(res)
        connection.commit()

def create_user(data):
    connection = sqlite3.connect("db_esign.db")
    cursor = connection.cursor()
    if data != "":
        res = f"INSERT INTO  User VALUES ({data})"
        cursor.execute(res)
        connection.commit()

def resend_mail(order_id):
    user = get_order('Create_by', 'Order_No', order_id)[0]
    mail = get_mail(user)
    time = get_order('Updated_Time', 'Order_No', order_id)[0]
    if order_id[:2] == "CO":
        send_mail(mail, f"result E-Sign  ({get_order('Order_Title', 'Order_No', order_id)[0]})",
                  f'the application have been signed!<br> at {time[0]} Please click to URL to see detail  http://10.220.40.75:5000/new_esign/{order_id}v')
    else:
        send_mail(mail, f"result E-Sign  ({get_order('Order_Title', 'Order_No', order_id)[0]})",
                  f'the application have been signed!<br> at {time[0]} Please click to URL to see detail  http://10.220.40.75:5000/esign/{order_id}')

def reject_mail(order_id):
    user = get_order('Create_by', 'Order_No', order_id)[0]
    mail = get_mail(user)
    if order_id[:2] == "CO":
        send_mail(mail, f"result E-Sign ({get_order('Order_Title', 'Order_No', order_id)[0]})",
                  'the application have been rejected!<br> Please click to URL to see detail  http://10.220.40.75:5000/new_search_order')
    else:
        send_mail(mail, f"result E-Sign ({get_order('Order_Title', 'Order_No', order_id)[0]})",
                  'the application have been rejected!<br> Please click to URL to see detail  http://10.220.40.75:5000/')

def check_signer(order_id, mail):
    ls_mail = get_cur_mail(order_id)
    status = get_order('Order_Status','Order_No',order_id)[0]
    if mail in ls_mail and status[:2] == 'Si':
        a = get_Signer_now(order_id)
        if mail in a.split(','):
            return True
        else:
            return False
    else:
        return False

def get_ls_co(mail):
    ls_order = []
    list_order = []
    orders = sql_ress(
        f"SELECT Order_No FROM Signature_Transaction WHERE  Order_Status NOT LIKE 'True' AND Order_Status NOT LIKE 'Reject'  AND File_Type = 'excel' ORDER BY Create_Time DESC")
    for i in range(10):
        order_id = sql_ress(
            f"SELECT Order_No FROM Signature_Transaction WHERE Signer_ID{i+1} LIKE '%{mail}%' AND Order_Status NOT LIKE 'True' AND Order_Status NOT LIKE 'Reject'  AND File_Type = 'excel' ORDER BY Create_Time ASC")

        if order_id != None:
            if len(order_id) == 1:
                temp = get_order(f'Status_Sign{i+1}', 'Order_No', order_id[0])
                if i+1 > 1 and temp[0] == "":
                    temp2 = get_order(
                        f'Status_Sign{i}', 'Order_No', order_id[0])
                    if temp2[0] != "" and temp2[0] != 'reject':
                        ls_order.append(order_id[0])
                else:
                    temp2 = get_order(
                        f'Status_Sign{i+1}', 'Order_No', order_id[0])
                    if temp2[0] == "":
                        ls_order.append(order_id[0])

            elif len(order_id) > 1:
                for id in order_id:
                    temp = get_order(f'Status_Sign{i+1}', 'Order_No', id)[0]
                    if i+1 > 1 and temp == "":
                        temp3 = get_order(f'Status_Sign{i}', 'Order_No', id)[0]
                        if temp3 != "" and temp3 != 'reject':
                            ls_order.append(id)
                    elif i+1 == 1 and temp == "":
                        ls_order.append(id)
    if orders != None:
        for k in orders:
            if k in ls_order and k[:2] == 'CO':
                list_order.append(k)
        return list_order
    return ls_order

def get_ls_ex(mail):
    ls_order = []
    list_order = []
    orders = sql_ress(f"SELECT Order_No FROM Signature_Transaction WHERE  Order_Status NOT LIKE 'True' AND Order_Status NOT LIKE 'Reject'  AND File_Type = 'Sign_excel' AND File_Type NOT LIKE 'MOQ/MPQ' ORDER BY Create_Time DESC")
    for i in range(10):
        order_id = sql_ress(
            f"SELECT Order_No FROM Signature_Transaction WHERE Signer_ID{i+1} LIKE '%{mail}%' AND Order_Status NOT LIKE 'True' AND Order_Status NOT LIKE 'Reject'  AND File_Type = 'Sign_excel' AND File_Type NOT LIKE 'MOQ/MPQ' ORDER BY Create_Time ASC")

        if order_id != None:
            if len(order_id) == 1:
                temp = get_order(f'Status_Sign{i+1}', 'Order_No', order_id[0])
                if i+1 > 1 and temp[0] == "":
                    temp2 = get_order(
                        f'Status_Sign{i}', 'Order_No', order_id[0])
                    if temp2[0] != "" and temp2[0] != 'reject':
                        ls_order.append(order_id[0])
                else:
                    temp2 = get_order(
                        f'Status_Sign{i+1}', 'Order_No', order_id[0])
                    if temp2[0] == "":
                        ls_order.append(order_id[0])

            elif len(order_id) > 1:
                for id in order_id:
                    temp = get_order(f'Status_Sign{i+1}', 'Order_No', id)[0]
                    if i+1 > 1 and temp == "":
                        temp3 = get_order(f'Status_Sign{i}', 'Order_No', id)[0]
                        if temp3 != "" and temp3 != 'reject':
                            ls_order.append(id)
                    elif i+1 == 1 and temp == "":
                        ls_order.append(id)
    for k in orders:
        if k in ls_order:
            list_order.append(k)

    return list_order

def get_sign_stt(order_id, param):
    qty = get_order('Signer_Qty', 'Order_No', order_id)[0]
    if param == 'reject':
        for i in range(qty):
            if get_order(f'Status_Sign{i+1}', 'Order_No', order_id)[0] == "reject":
                reject_by = get_order(
                    f'Signer_ID{i+1}', 'Order_No', order_id)[0]
                return reject_by, get_ID_mail(reject_by), get_order('Update_Time', 'Order_No', order_id)[0]

def get_Signer_now(id):
    qty = get_order('Signer_Qty', 'Order_No', id)[0]
    a = ""
    for i in range(qty):
        if get_order(f'Status_Sign{i+1}', 'Order_No', id)[0] == "":
            if get_order(f'Signer_ID{i+1}', 'Order_No', id)[0].find(",") == -1:
                return get_order(f'Signer_ID{i+1}', 'Order_No', id)[0]
            else:
                for i, mail in enumerate(get_order(f'Signer_ID{i+1}', 'Order_No', id)[0].split(",")):
                    if i == (len(get_order(f'Signer_ID{i+1}', 'Order_No', id)[0].split(","))-1):
                        a += mail
                    else:
                        a += mail + ","
                return a
        else:
            if get_order(f'Status_Sign{i+1}', 'Order_No', id)[0] == "reject":
                if get_order(f'Signer_ID{i+1}', 'Order_No', id)[0].find(",") == -1:
                    return get_order(f'Signer_ID{i+1}', 'Order_No', id)[0]
                else:
                    for i, mail in enumerate(get_order(f'Signer_ID{i+1}', 'Order_No', id)[0].split(",")):
                        if i == (len(get_order(f'Signer_ID{i+1}', 'Order_No', id)[0].split(","))-1):
                            a += mail
                        else:
                            a += mail + ","
                    return a
            elif get_order(f'Status_Sign{i+1}', 'Order_No', id)[0] == "approve" and get_order('Order_Status', 'Order_No', id)[0] == 'True':
                if get_order(f'Signer_ID{i+1}', 'Order_No', id)[0].find(",") == -1:
                    return get_order(f'Signer_ID{i+1}', 'Order_No', id)[0]
                else:
                    for i, mail in enumerate(get_order(f'Signer_ID{i+1}', 'Order_No', id)[0].split(",")):
                        if i == (len(get_order(f'Signer_ID{i+1}', 'Order_No', id)[0].split(","))-1):
                            a += mail

                        else:
                            a += mail + ","
                    return a
            elif get_order(f'Status_Sign{i+1}', 'Order_No', id)[0] != "":
                if get_order(f'Signer_ID{i+2}', 'Order_No', id)[0].find(",") == -1:
                    return get_order(f'Signer_ID{i+2}', 'Order_No', id)[0]
                else:
                    for i, mail in enumerate(get_order(f'Signer_ID{i+2}', 'Order_No', id)[0].split(",")):
                        if i == (len(get_order(f'Signer_ID{i+2}', 'Order_No', id)[0].split(","))-1):
                            a += mail
                        else:
                            a += mail + ","
                    return a
            else:
                return ""

def get_ls_order_sign(mail):
    list_order = []
    ls_orders = []
    
    for a in range(10):
        id_order = sql_ress(
            f"SELECT Order_No FROM Signature_Transaction WHERE Signer_ID{a+1} LIKE '%{mail}%' AND Order_Status NOT LIKE 'True' AND Order_No NOT LIKE 'CO%' AND Order_Status NOT LIKE 'Reject'  AND File_Type NOT LIKE 'Sign_excel' AND File_Type NOT LIKE 'MOQ/MPQ' ORDER BY Create_Time ASC")
        print(id_order)
        if id_order != None:
            if len(id_order) == 1:
                temp = get_order(
                    f'Status_Sign{a+1}', 'Order_No', id_order[0])[0]
                if (a+1) > 1 and temp == "":
                    temp2 = get_order(
                        f'Status_Sign{a}', 'Order_No', id_order[0])[0]
                    if temp2 != "" and temp2 != 'reject':
                        list_order.append(id_order[0])
                else:
                    temp2 = get_order(
                        f'Status_Sign{a+1}', 'Order_No', id_order[0])[0]
                    if temp2 == "":
                        list_order.append(id_order[0])
            elif len(id_order) > 1:
                for id in id_order:
                    temp = get_order(f'Status_Sign{a+1}', 'Order_No', id)[0]
                    if (a+1) > 1 and temp == "":
                        temp3 = get_order(f'Status_Sign{a}', 'Order_No', id)[0]
                        if temp3[0] != "" and temp3[0] != 'reject':
                            list_order.append(id)
                    elif (a+1) == 1 and temp == "":
                        list_order.append(id)
    orders = sql_ress(f"SELECT Order_No FROM Signature_Transaction WHERE Order_Status NOT LIKE 'True' AND Order_Status NOT LIKE 'Reject'  AND File_Type NOT LIKE 'Sign_excel' AND File_Type NOT LIKE 'MOQ/MPQ' ORDER BY Create_Time DESC")
    if orders:
        for i in orders:
            if i in list_order:
                ls_orders.append(i)
        return ls_orders
    else:
        return ''

def sign_ok(username, mail, order_id, result, loc, reason, note):
    today = datetime.now()
    time = today.strftime("%m/%d/%Y %H:%M:%S")
    cks = get_order(f'Signer_ID{loc}', 'Order_No', order_id)[0]
    status = get_order('Order_Status', 'Order_No', order_id)[0]
    title = get_order('Order_Title', 'Order_No', order_id)[0]
    if cks.find(mail) != -1 and status != 'Reject':
        qty = get_order('Signer_Qty', 'Order_No', order_id)[0]
        if loc < int(qty):
            res = f"UPDATE Signature_Transaction SET Signer_ID{loc} = '{mail}',Time_Signer{loc} = '{time}',  Status_Sign{loc} = '{result}', Updated_Time = '{time}', Updated_By = '{username}' WHERE Order_No = '{order_id}'"
            update_order(res)
            if result == 'approve':
                next_mail = get_order(
                    f'Signer_ID{loc+1}', 'Order_No', order_id)[0]
                res = f"UPDATE Signature_Transaction SET Order_Status = 'Signed: {loc}/{qty}' WHERE Order_No = '{order_id}'"
                update_order(res)
                create_by = get_create_by(order_id)[0]
                if order_id[:2] == "CO":
                    send_mail(
                        next_mail, title, f'Dear User!<br> You have 1 Order waiting to be signed!<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)} <br>  Please click to this URL to Sign order http://10.220.40.75:5000/new_esign/{order_id}s <br> Note: {note}')
                    resend_mail(order_id)
                else:
                    send_mail(
                        next_mail, title, f'Dear User!<br> You have 1 Order waiting to be signed!<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)} <br>  Please click to this URL to Sign order http://10.220.40.75:5000/esign/{order_id} <br> Note: {note}')
                    resend_mail(order_id)
            else:
                res = f"UPDATE Signature_Transaction SET Signer_ID{loc} = '{mail}',Time_Signer{loc} = '{time}',  Status_Sign{loc} = '{result}', Order_Status = 'Reject' , Updated_Time = '{time}', Updated_By = '{username}', Reason = '{reason}' WHERE Order_No = '{order_id}'"
                update_order(res)
                reject_mail(order_id)
        elif loc == int(qty):
            if result == 'approve':
                res = f"UPDATE Signature_Transaction SET Signer_ID{loc} = '{mail}',Time_Signer{loc} = '{time}',  Status_Sign{loc} = '{result}', Order_Status = 'True' , Updated_Time = '{time}', Updated_By = '{username}' WHERE Order_No = '{order_id}'"
                update_order(res)
                resend_mail(order_id)
            else:
                res = f"UPDATE Signature_Transaction SET Signer_ID{loc} = '{mail}',Time_Signer{loc} = '{time}',  Status_Sign{loc} = '{result}', Order_Status = 'Reject' , Updated_Time = '{time}', Updated_By = '{username}', Reason = '{reason}' WHERE Order_No = '{order_id}'"
                update_order(res)
                reject_mail(order_id)

def mpm_sign_ok(username, mail, order_id, result, loc, reason):
    today = datetime.now()
    time = today.strftime("%m/%d/%Y %H:%M:%S")
    cks = get_order(f'Signer_ID{loc}', 'Order_No', order_id)[0]
    status = get_order('Order_Status', 'Order_No', order_id)[0]
    title = get_order('Order_Title', 'Order_No', order_id)[0]
    if cks == mail and status != 'Reject':
        qty = qty = get_order('Signer_Qty', 'Order_No', order_id)[0]
        if loc < int(qty):
            res = f"UPDATE Signature_Transaction SET Time_Signer{loc} = '{time}',  Status_Sign{loc} = '{result}', Updated_Time = '{time}', Updated_By = '{username}' WHERE Order_No = '{order_id}'"
            update_order(res)
            if result == 'approve':
                next_mail = get_order(
                    f'Signer_ID{loc+1}', 'Order_No', order_id)[0]
                name = get_Name_mail(next_mail)
                res = f"UPDATE Signature_Transaction SET Order_Status = 'Signed: {loc}/{qty}' WHERE Order_No = '{order_id}'"
                update_order(res)
                create_by = get_create_by(order_id)[0]
                send_mail(
                    next_mail, title, f'Dear User!<br> You have 1 Order waiting to be signed!<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)}<br>   Please click to this URL to Sign order http://10.220.40.75:5000/MPM/esign/{order_id}')
            else:
                res = f"UPDATE Signature_Transaction SET Time_Signer{loc} = '{time}',  Status_Sign{loc} = '{result}', Order_Status = 'Reject' , Updated_Time = '{time}', Updated_By = '{username}', Reason = '{reason}' WHERE Order_No = '{order_id}'"
                update_order(res)
                reject_mail(order_id)
        elif loc == int(qty):
            if result == 'approve':
                res = f"UPDATE Signature_Transaction SET Time_Signer{loc} = '{time}',  Status_Sign{loc} = '{result}', Order_Status = 'True' , Updated_Time = '{time}', Updated_By = '{username}' WHERE Order_No = '{order_id}'"
                update_order(res)
                resend_mail(order_id)
            else:
                res = f"UPDATE Signature_Transaction SET Time_Signer{loc} = '{time}',  Status_Sign{loc} = '{result}', Order_Status = 'Reject' , Updated_Time = '{time}', Updated_By = '{username}', Reason = '{reason}' WHERE Order_No = '{order_id}'"
                update_order(res)
                reject_mail(order_id)

def get_mpm_order_sign(mail):
    list_order = []
    for a in range(10):
        id_order = sql_ress(
            f"SELECT Order_No FROM Signature_Transaction WHERE Signer_ID{a+1} = '{mail}' AND Order_Status NOT LIKE 'True' AND Order_Status NOT LIKE 'Reject'  AND File_Type LIKE 'MOQ/MPQ' ORDER BY Create_Time ASC")
        if id_order != None:
            if len(id_order) == 1:
                temp = get_order(f'Status_Sign{a+1}', 'Order_No', id_order[0])
                if (a+1) > 1 and temp[0] == "":
                    temp2 = get_order(
                        f'Status_Sign{a}', 'Order_No', id_order[0])
                    if temp2[0] != "" and temp2[0] != 'reject':
                        list_order.append(id_order[0])
                else:
                    temp2 = get_order(
                        f'Status_Sign{a+1}', 'Order_No', id_order[0])
                    if temp2[0] == "":
                        list_order.append(id_order[0])
            elif len(id_order) > 1:
                for id in id_order:
                    temp = get_order(f'Status_Sign{a+1}', 'Order_No', id)[0]
                    if (a+1) > 1 and temp == "":
                        temp3 = get_order(f'Status_Sign{a}', 'Order_No', id)[0]
                        if temp3 != "" and temp3 != 'reject':
                            list_order.append(id)
                    elif (a+1) == 1 and temp == "":
                        list_order.append(id)
    return list_order

def get_cur_mail(order_id):
    a = ""
    cur_mails = []
    create_id = get_order('Create_By', 'Order_no', order_id)[0]
    a += get_mail(create_id)+","
    for i in range(10):
        a += sql_ress(
            f"SELECT Signer_ID{i+1} FROM Signature_Transaction WHERE Order_No = '{order_id}'")[0] + ","
    for mail in a[:-1].split(","):
        if mail != "":
            cur_mails.append(mail)
    return cur_mails
# =========================================END Function============================================

# =========================================Signature===============================================
@app.route('/sign_normal', methods=['GET', 'POST'])
def main():
    ls_createby = []
    ls_createtime = []
    ls_title = []
    if not session.get("username"):
        session["url"] = "main"
        return redirect(url_for('login'))
    else:
        session["url"] = "main"
        username = session.get("username")
        mail = get_mail(username)
        ls_order = get_ls_order_sign(mail)
        for id in ls_order:
            ls_createby.append(get_order('Create_By', 'Order_No', id)[0])
            ls_createtime.append(get_order('Create_Time', 'Order_No', id)[0])
            ls_title.append(get_order('Order_Title', 'Order_No', id)[0])
    return render_template('main.html', ls_order=ls_order, ls_createby=ls_createby, ls_createtime=ls_createtime, ls_title=ls_title, task = 'normal')

@app.route('/sign_excel', methods=['GET', 'POST'])
def sign_excel():
    ls_createby = []
    ls_createtime = []
    ls_title = []
    new_ls = []
    if not session.get("username"):
        session["url"] = "sign_excel"
        return redirect(url_for('login'))
    else:
        session["url"] = "sign_excel"
        username = session.get("username")
        mail = get_mail(username)
        ls_order = get_ls_ex(mail)
        for id in ls_order:
            if id[:2] != 'SO' and id[:2] != 'De' and id[:2] != 'CO':
                new_ls.append[id]
                ls_createby.append(get_order('Create_By', 'Order_No', id)[0])
                ls_createtime.append(get_order('Create_Time', 'Order_No', id)[0])
                ls_title.append(get_order('Order_Title', 'Order_No', id)[0])
    return render_template('main.html', ls_order=new_ls, ls_createby=ls_createby, ls_createtime=ls_createtime, ls_title=ls_title,task = 'excel')

@app.route('/edit/<order_id>', methods=['POST', 'GET'])
def edit(order_id):
    ls_signer = []
    ls_time = []
    ls_status = []
    ls_table = []
    ls_file2 = ""
    file_type = get_order('File_Type', 'Order_No', order_id)[0]
    ls_file = get_order('File_name', 'Order_No', order_id)[0]
    for i in range(get_order('Signer_Qty', 'Order_No', order_id)[0]):
        ls_signer.append(get_order(f'Signer_ID{i+1}', 'Order_No', order_id)[0])
        ls_time.append(get_order(f'Time_Signer{i+1}', 'Order_No', order_id)[0])
        ls_status.append(
            get_order(f'Status_Sign{i+1}', 'Order_No', order_id)[0])
    username = get_order('Create_By', 'Order_No', order_id)[0]
    if get_order('Order_Status', 'Order_No', order_id)[0] == 'Reject':
        status = 'Rejected'
    elif get_order('Order_Status', 'Order_No', order_id)[0] == 'True':
        status = 'Signed'
    else:
        status = 'Waiting'
    context = {
        'employee': username, 'order_id': order_id, 'ename': get_Name_ID(username), 'title': get_order('Order_Title', 'Order_No', order_id)[0], 'content': get_order('Description', 'Order_No', order_id)[0], 'status': status, 'reason': get_order('Reason', 'Order_No', order_id)[0], 'create_at': get_order('Create_Time', 'Order_No', order_id)[0]
    }

    if request.method == 'POST':
        new_content = request.form['content']
        loc = 0
        qty = get_order('Signer_Qty', 'Order_No', order_id)[0]
        for i in range(10):
            if get_order(f'Status_Sign{i+1}', 'Order_No', order_id)[0] == "reject":
                loc = (i+1)
        if file_type == "Sign_excel":
            file = request.files['re_file']
            if file.filename != "":
                uploads_dir = os.path.join(
                    f"data/{username.upper()}/Excel_Sign", order_id)
                os.makedirs(uploads_dir, exist_ok=True)
                objects = os.listdir(uploads_dir)
                files_file = [f for f in objects if os.path.isfile(
                    os.path.join(uploads_dir, f))]
                for f in files_file:
                    os.remove(os.path.join(uploads_dir, f))
                file.save(os.path.join(uploads_dir, file.filename))
                res = f"UPDATE Signature_Transaction SET Order_Status = 'Signed: {loc-1}/{qty}', Time_Signer{loc} = '', Status_Sign{loc} = '', Description = '{new_content}', File_name = '{file.filename}', Reason = '' WHERE Order_No = '{order_id}'"
                update_order(res)
            else:
                file_name = get_order('File_name', 'Order_No', order_id)[0]
                uploads_dir = os.path.join(
                    f"data/{username.upper()}/Excel_Sign", order_id)
                shutil.copy(uploads_dir+"/result.xls",
                            uploads_dir+"/"+file_name)
                res = f"UPDATE Signature_Transaction SET Order_Status = 'Signed: {loc-1}/{qty}', Time_Signer{loc} = '', Status_Sign{loc} = '', Description = '{new_content}', Reason = '' WHERE Order_No = '{order_id}'"
                update_order(res)
        elif file_type == "excel":
            files = request.files.getlist('re_files')

            if files[0].filename != "":
                uploads_dir = os.path.join(
                    f"data/{username.upper()}/Normal_Sign", order_id)
                os.makedirs(uploads_dir, exist_ok=True)
                objects = os.listdir(uploads_dir)
                files_file = [f for f in objects if os.path.isfile(
                    os.path.join(uploads_dir, f))]
                for f in files_file:
                    os.remove(os.path.join(uploads_dir, f))
                for file in request.files.getlist('re_files'):
                    if file.filename != "":
                        file.save(os.path.join(uploads_dir, file.filename))
                        ls_file2 += file.filename+","
                res = f"UPDATE Signature_Transaction SET Order_Status = 'Signed: {loc-1}/{qty}', Time_Signer{loc} = '', Status_Sign{loc} = '', Description = '{new_content}', File_name = '{ls_file2}', Reason = '' WHERE Order_No = '{order_id}'"
                update_order(res)
            else:
                res = f"UPDATE Signature_Transaction SET Order_Status = 'Signed: {loc-1}/{qty}', Time_Signer{loc} = '', Status_Sign{loc} = '', Description = '{new_content}', Reason = '' WHERE Order_No = '{order_id}'"
                update_order(res)

        re_mail = get_order(f'Signer_ID{loc}', 'Order_No', order_id)[0]
        create_by = get_create_by(order_id)[0]
        send_mail(re_mail, get_order('Order_Title', 'Order_No', order_id)[
                  0], f'Dear User!<br>  You have 1 Order waiting to be signed!<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)}<br>  Please click to this URL to Sign order http://10.220.40.75:5000/esign/{order_id}')
        return redirect(url_for('search_order'))
    if ls_file != "":
        if file_type == "excel":
            uploads_dir = os.path.join(
                f"data/{username.upper()}/Normal_Sign", order_id)
            ls_file = ls_file.split(",")
            if ls_file[-1] == "":
                ls_file = ls_file[:-1]
            else:
                ls_file = ls_file
            for file in ls_file:
                wb = pd.read_excel(uploads_dir+"/"+file)
                wb = wb.dropna(how='all')
                wb = wb.fillna("")
                wb = wb.reset_index(drop=True)
                ls_table.append(wb.to_html(index=False, classes='normal_sign mb-0 table table-bordered'))
        elif file_type == "Sign_excel":
            uploads_dir = os.path.join(
                f"data/{username.upper()}/Excel_Sign", order_id)
            wb = pd.read_excel(uploads_dir+"/result.xls")
            wb = wb.dropna(how='all')
            wb = wb.fillna("")
            wb = wb.reset_index(drop=True)
            ls_table.append(wb.to_html(
                index=False, classes='excel_sign mb-0 table table-bordered', table_id='excel_sign'))

        return render_template('edit.html', ls_signer=ls_signer, context=context, ls_file=ls_file, ls_time=ls_time, ls_status=ls_status, tables=ls_table, file_type=file_type)
    else:
        return render_template('edit.html', ls_signer=ls_signer, context=context, ls_file=ls_file, ls_time=ls_time, ls_status=ls_status, file_type=file_type)

@app.route('/get_all', methods=['GET', 'POST'])
def get_all():
    ls_stt = []
    ls_Sign = []
    ls_createtime = []
    ls_signer = []
    ls_title = []
    ls_updatetime = []
    if not session.get("username"):
        session["url"] = 'get_all'
        return redirect(url_for('login'))
    else:
        session["url"] = 'get_all'
        username = session.get("username")

        list_order = sql_ress(f"SELECT Order_No FROM Signature_Transaction WHERE Create_By like '%{username}%' AND Order_No not like 'SO%' AND Order_No not like 'CO%' AND Order_No not like 'De%' ORDER BY Updated_Time DESC")
        print(list_order)
        if request.method == 'POST':
            task = request.form['task']
            alert_id = request.form['order_id']
            mail = get_Signer_now(alert_id)
            s_mail = get_mail(username)
            ls_title = get_order('Order_Title', 'Order_No', alert_id)[0]
            create_by = get_create_by(alert_id)[0]
            if task == 'alert':
                send_mail(mail, "(Alert!!!) " + ls_title,
                          f'Dear User!<br>  You have 1 Order waiting to be signed!<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)}<br> Please click to this URL to Sign order http://10.220.40.75:5000/esign/{alert_id}')
                send_mail(s_mail, "(Alert!!!) " + ls_title,
                          f'Dear User!<br>  Your Order waiting to be signed! <br> Please Contact {alert_id} -- {mail} to sign the application')
                return 'done'

        if list_order != None:
            for i in list_order:
                status = get_order('Order_Status', 'Order_No', i)[0]
                if status == "Reject":
                    ls_Sign.append(get_Signer_now(i))
                    ls_stt.append("Rejected!!!")
                elif status == "True":
                    ls_stt.append("Signed!!!")
                    ls_Sign.append(get_Signer_now(i))
                else:
                    ls_stt.append(f"{status}")
                    ls_Sign.append(get_Signer_now(i))

                ls_updatetime.append(
                    get_order('Updated_Time', 'Order_No', i)[0])
                ls_createtime.append(
                    get_order('Create_Time', 'Order_No', i)[0])
                ls_title.append(get_order('Order_Title', 'Order_No', i)[0])

            for sign in ls_Sign:
                ls_signer.append(get_ID_mail(sign))
            return render_template('search_order.html', ls_updatetime=ls_updatetime, list_order=list_order, ls_stt=ls_stt, ls_Sign=ls_Sign, ls_createtime=ls_createtime, ls_signer=ls_signer, ls_title=ls_title, task = 'all')
        else:
            return render_template('search_order.html')

@app.route('/get_waiting', methods=['GET', 'POST'])
def get_waiting():
    ls_stt = []
    ls_Sign = []
    ls_createtime = []
    ls_signer = []
    ls_title = []
    ls_updatetime = []
    if not session.get("username"):
        session["url"] = 'get_waiting'
        return redirect(url_for('login'))
    else:
        session["url"] = 'get_waiting'
        username = session.get("username")

        list_order = sql_ress(
            f"SELECT Order_No FROM Signature_Transaction WHERE Create_By like '%{username}%' AND Order_No not like 'SO%' AND Order_No not like 'CO%' AND Order_No not like 'De%' ORDER BY Updated_Time DESC")
        if request.method == 'POST':
            task = request.form['task']
            alert_id = request.form['order_id']
            mail = get_Signer_now(alert_id)
            s_mail = get_mail(username)
            ls_title = get_order('Order_Title', 'Order_No', alert_id)[0]
            create_by = get_create_by(alert_id)[0]
            if task == 'alert':
                send_mail(mail, "(Alert!!!) " + ls_title,
                          f'Dear User!<br>  You have 1 Order waiting to be signed!<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)} <br> Please click to this URL to Sign order http://10.220.40.75:5000/esign/{alert_id}')
                send_mail(s_mail, "(Alert!!!) " + ls_title,
                          f'Dear User!<br>  Your Order waiting to be signed! <br> Please Contact {alert_id} -- {mail} to sign the application')
                return 'done'

        if list_order != None:
            for i in list_order:
                status = get_order('Order_Status', 'Order_No', i)[0]
                if status == "Reject":
                    ls_Sign.append(get_Signer_now(i))
                    ls_stt.append("Rejected!!!")
                elif status == "True":
                    ls_stt.append("Signed!!!")
                    ls_Sign.append(get_Signer_now(i))
                else:
                    ls_stt.append(f"{status}")
                    ls_Sign.append(get_Signer_now(i))

                ls_updatetime.append(
                    get_order('Updated_Time', 'Order_No', i)[0])
                ls_createtime.append(
                    get_order('Create_Time', 'Order_No', i)[0])
                ls_title.append(get_order('Order_Title', 'Order_No', i)[0])

            for sign in ls_Sign:
                ls_signer.append(get_ID_mail(sign))
            return render_template('search_order.html', ls_updatetime=ls_updatetime, list_order=list_order, ls_stt=ls_stt, ls_Sign=ls_Sign, ls_createtime=ls_createtime, ls_signer=ls_signer, ls_title=ls_title, task = 'waiting')
        else:
            return render_template('search_order.html')
    
@app.route('/get_signed', methods=['GET', 'POST'])
def get_signed():
    ls_stt = []
    ls_Sign = []
    ls_createtime = []
    ls_signer = []
    ls_title = []
    ls_updatetime = []
    if not session.get("username"):
        session["url"] = 'get_signed'
        return redirect(url_for('login'))
    else:
        session["url"] = 'get_signed'
        username = session.get("username")

        list_order = sql_ress(
            f"SELECT Order_No FROM Signature_Transaction WHERE Create_By like '%{username}%' AND Order_No not like 'SO%' AND Order_No not like 'CO%' AND Order_No not like 'De%' ORDER BY Updated_Time DESC")
        if request.method == 'POST':
            task = request.form['task']
            alert_id = request.form['order_id']
            mail = get_Signer_now(alert_id)
            s_mail = get_mail(username)
            ls_title = get_order('Order_Title', 'Order_No', alert_id)[0]
            create_by = get_create_by(alert_id)[0]
            if task == 'alert':
                send_mail(mail, "(Alert!!!) " + ls_title,
                          f'Dear User!<br>  You have 1 Order waiting to be signed!<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)} <br> Please click to this URL to Sign order http://10.220.40.75:5000/esign/{alert_id}')
                send_mail(s_mail, "(Alert!!!) " + ls_title,
                          f'Dear User!<br>  Your Order waiting to be signed! <br> Please Contact {alert_id} -- {mail} to sign the application')
                return 'done'

        if list_order != None:
            for i in list_order:
                status = get_order('Order_Status', 'Order_No', i)[0]
                if status == "Reject":
                    ls_Sign.append(get_Signer_now(i))
                    ls_stt.append("Rejected!!!")
                elif status == "True":
                    ls_stt.append("Signed!!!")
                    ls_Sign.append(get_Signer_now(i))
                else:
                    ls_stt.append(f"{status}")
                    ls_Sign.append(get_Signer_now(i))

                ls_updatetime.append(
                    get_order('Updated_Time', 'Order_No', i)[0])
                ls_createtime.append(
                    get_order('Create_Time', 'Order_No', i)[0])
                ls_title.append(get_order('Order_Title', 'Order_No', i)[0])

            for sign in ls_Sign:
                ls_signer.append(get_ID_mail(sign))
            return render_template('search_order.html', ls_updatetime=ls_updatetime, list_order=list_order, ls_stt=ls_stt, ls_Sign=ls_Sign, ls_createtime=ls_createtime, ls_signer=ls_signer, ls_title=ls_title, task = 'signed')
        else:
            return render_template('search_order.html')

@app.route('/get_rejected', methods=['GET', 'POST'])
def get_rejected():
    ls_stt = []
    ls_Sign = []
    ls_createtime = []
    ls_signer = []
    ls_title = []
    ls_updatetime = []
    if not session.get("username"):
        session["url"] = 'get_rejected'
        return redirect(url_for('login'))
    else:
        session["url"] = 'get_rejected'
        username = session.get("username")

        list_order = sql_ress(
            f"SELECT Order_No FROM Signature_Transaction WHERE Create_By like '%{username}%' AND Order_No not like 'SO%' AND Order_No not like 'CO%' AND Order_No not like 'De%' ORDER BY Updated_Time DESC")
        if request.method == 'POST':
            task = request.form['task']
            alert_id = request.form['order_id']
            mail = get_Signer_now(alert_id)
            s_mail = get_mail(username)
            ls_title = get_order('Order_Title', 'Order_No', alert_id)[0]
            create_by = get_create_by(alert_id)[0]
            if task == 'alert':
                send_mail(mail, "(Alert!!!) " + ls_title,
                          f'Dear User!<br>  You have 1 Order waiting to be signed!<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)} <br> Please click to this URL to Sign order http://10.220.40.75:5000/esign/{alert_id}')
                send_mail(s_mail, "(Alert!!!) " + ls_title,
                          f'Dear User!<br>  Your Order waiting to be signed! <br> Please Contact {alert_id} -- {mail} to sign the application')
                return 'done'

        if list_order != None:
            for i in list_order:
                status = get_order('Order_Status', 'Order_No', i)[0]
                if status == "Reject":
                    ls_Sign.append(get_Signer_now(i))
                    ls_stt.append("Rejected!!!")
                elif status == "True":
                    ls_stt.append("Signed!!!")
                    ls_Sign.append(get_Signer_now(i))
                else:
                    ls_stt.append(f"{status}")
                    ls_Sign.append(get_Signer_now(i))

                ls_updatetime.append(
                    get_order('Updated_Time', 'Order_No', i)[0])
                ls_createtime.append(
                    get_order('Create_Time', 'Order_No', i)[0])
                ls_title.append(get_order('Order_Title', 'Order_No', i)[0])

            for sign in ls_Sign:
                ls_signer.append(get_ID_mail(sign))
            return render_template('search_order.html', ls_updatetime=ls_updatetime, list_order=list_order, ls_stt=ls_stt, ls_Sign=ls_Sign, ls_createtime=ls_createtime, ls_signer=ls_signer, ls_title=ls_title, task = 'reject')
        else:
            return render_template('search_order.html')

@app.route('/detail/<order_id>', methods=['GET', 'POST'])
def detail(order_id):
    ls_signer = []
    ls_file = []
    ls_time = []
    ls_status = []
    ls_table = []
    order_id = order_id[:-1]
    username = get_order('Create_By', 'Order_No', order_id)[0]
    file_style = get_order('File_Type', 'Order_No', order_id)[0]
    if file_style != "Sign_excel":
        uploads_dir = os.path.join(
            f"data/{username.upper()}/Normal_Sign", order_id)
    else:
        uploads_dir = os.path.join(
            f"data/{username.upper()}/Excel_Sign", order_id)
    ls_file = get_order('File_name', 'Order_No', order_id)[0]
    hig_line = get_order('File_Link', 'Order_No', order_id)[0]
    reason = get_order('Reason', 'Order_No', order_id)[0]

    if get_order('Order_Status', 'Order_No', order_id)[0] == 'Reject':
        status = 'Rejected'
    elif get_order('Order_Status', 'Order_No', order_id)[0] == 'True':
        status = 'Signed'
    else:
        status = 'Waiting Sign'
    if ls_file != None:
        ls_file = ls_file.split(",")
        if file_style != "Sign_excel":
            if ls_file[-1] == "":
                ls_file = ls_file[:-1]
            else:
                ls_file = ls_file
            for file in ls_file:
                wb = pd.read_excel(uploads_dir+"/"+file)
                wb = wb.dropna(how='all')
                wb = wb.fillna("")
                wb = wb.reset_index(drop=True)
                ls_table.append(wb.to_html(
                    index=False, classes='normal_sign tbl1 mb-0 table table-bordered'))
        else:
            if not os.path.exists(uploads_dir+"/result.xls"):
                wb = pd.read_excel(uploads_dir+"/"+ls_file[0])
            else:
                wb = pd.read_excel(uploads_dir+"/result.xls")
            wb = wb.dropna(how='all')
            wb = wb.fillna("")
            wb = wb.reset_index(drop=True)

            ls_table.append(wb.to_html(
                index=False, classes='excel_sign tbl1 mb-0 table table-bordered', table_id='excel_sign'))
    for i in range(get_order('Signer_Qty', 'Order_No', order_id)[0]):
        ls_signer.append(get_order(f'Signer_ID{i+1}', 'Order_No', order_id)[0])
        ls_time.append(get_order(f'Time_Signer{i+1}', 'Order_No', order_id)[0])
        ls_status.append(
            get_order(f'Status_Sign{i+1}', 'Order_No', order_id)[0])
    context = {

        'reason': reason, 'hig_line': hig_line, 'employee': username, 'order_id': order_id, 'ename': get_Name_ID(username), 'title': get_order('Order_Title', 'Order_No', order_id)[0], 'content': get_order('Description', 'Order_No', order_id)[0], 'status': status, 'create_at': get_order('Create_Time', 'Order_No', order_id)[0]
    }

    if ls_file != None:
        return render_template('order_detail.html', tables=ls_table, ls_signer=ls_signer, context=context, ls_file=ls_file, ls_time=ls_time, ls_status=ls_status)
    else:
        return render_template('order_detail.html', ls_signer=ls_signer, context=context, ls_time=ls_time, ls_status=ls_status)

@app.route('/esign/<order_id>', methods=['GET', 'POST'])
def esign(order_id):
    session["order_id"] = order_id
    ls_table = []
    create_by = get_order('Create_By', 'Order_No', order_id)[0]
    ls_file = get_order('File_name', 'Order_No', order_id)[0]
    file_style = get_order('File_Type', 'Order_No', order_id)[0]
   
    if not session.get("username"):
        session["url"] = "esign"
        return redirect(url_for('login'))
    else:
        session.pop("order_id")
        username = session.get("username")
        mail = get_mail(username)
        print(username)
        print("000000000000000000000",check_signer(order_id, mail))
        if check_signer(order_id, mail):
            for a in range(10):
                temp = sql_ress(
                    f"SELECT Status_Sign{a+1} FROM Signature_Transaction WHERE Signer_ID{a+1} LIKE '%{mail}%' AND Order_No = '{order_id}'")
                if temp != None:
                    if temp[0] == "" and (a+1) > 1:
                        temp2 = get_order(
                            f'Status_Sign{a}', 'Order_No', order_id)[0]
                        if temp2 != "":
                            loc = a+1
                        else:
                            loc = a
                        session["loc"] = loc
                    elif temp[0] == "" and (a+1) == 1:
                        session["loc"] = 1
        else:
            return render_template('signed.html')
        if request.method == 'POST':
            task = request.form['task']
            note = request.form['note']
            if task == '':
                if file_style != "Sign_excel":
                    uploads_dir = os.path.join(
                        f"data/{create_by.upper()}/Normal_Sign", order_id)
                    lc = session.get("loc")
                    reason = request.form['reason']
                    result = request.form['result']
                    sign_ok(username, mail, order_id, result, lc, reason, note)
                    return 'Done'
                else:
                    reason = request.form['reason2']
                    uploads_dir = os.path.join(
                        f"data/{create_by.upper()}/Excel_Sign", order_id)
                    result_sign = request.form['result']
                    lc = session.get("loc")
                    new_sign = ""
                    ls_index = get_order('File_Link', 'Order_No', order_id)[0]
                    if ls_index == "":
                        res = f"UPDATE Signature_Transaction SET File_Link = '{result_sign}' WHERE Order_No = '{order_id}'"
                        update_order(res)
                    else:

                        ls_index = ls_index.split(",")
                        for i in result_sign.split(",")[:-1]:
                            new_sign += str(ls_index[int(i)])+","

                        res = f"UPDATE Signature_Transaction SET File_Link = '{new_sign}' WHERE Order_No = '{order_id}'"
                        update_order(res)
                    if not os.path.exists(uploads_dir+"/result.xls"):
                        shutil.copy(uploads_dir+"/"+ls_file.split(",")
                                    [0], uploads_dir+"/result.xls")
                    wb = pd.read_excel(uploads_dir+"/"+ls_file.split(",")[0])
                    result_sign = result_sign.split(",")[:-1]
                    if order_id[:2] == "SO":
                        if(len(result_sign) > 1):
                            result = 'approve'
                        else:
                            result = 'reject'
                    else:
                        if(len(result_sign) != 0):
                            result = 'approve'
                        else:
                            result = 'reject'
                    a = wb.iloc[result_sign]
                    da = pd.DataFrame(a)

                    da.to_excel(uploads_dir+"/" +
                                ls_file.split(",")[0], index=None)
                    sign_ok(username, mail, order_id, result, lc, reason, note)
                    return 'Done'
            else:
                lc = session.get("loc")
                new_mail = request.form['new_mail']
                if get_order(f'Signer_ID{lc}', 'Order_No', order_id)[0].find(",") == -1:
                    signers = get_order(
                        f'Signer_ID{lc}', 'Order_No', order_id)[0]
                    res = f"UPDATE Signature_Transaction SET Signer_ID{lc} = '{new_mail}' WHERE Order_No = '{order_id}'"
                    create_by = get_create_by(order_id)[0]
                    update_order(res)
                    send_mail(
                        new_mail, order_id, f"Dear User!<br> Now you becoming the Signer of the application {order_id}<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)}<br> Please click to this URL to Sign order http://10.220.40.75:5000/esign/{order_id}")
                else:
                    signers = get_order(f'Signer_ID{lc}', 'Order_No', order_id)[
                        0].split(",")
                    a = ""
                    for i in signers:
                        if i != "" and i == mail:
                            a += new_mail + ","
                        else:
                            a += i + ","
                    res = f"UPDATE Signature_Transaction SET Signer_ID{lc} = '{a[:-2]}' WHERE Order_No = '{order_id}'"
                    update_order(res)
                    send_mail(
                        new_mail, order_id, f"Dear User!<br> Now you becoming the Signer of the application {order_id}<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)}<br> Please click to this URL to Sign order http://10.220.40.75:5000/esign/{order_id}")
                return 'Done'
        context = {
            'fullname': get_Name_ID(create_by), 'title': get_order('Order_Title', 'Order_No', order_id)[0], 'content': get_order('Description', 'Order_No', order_id)[0], 'order_id': order_id, 'create_by': create_by, 'file_type': file_style, 'mail': get_mail(create_by), 'cur_mail': get_cur_mail(order_id)
        }

        # qty = get_oder('Signer_Qty','')
        if ls_file != "":
            if file_style != "Sign_excel":
                uploads_dir = os.path.join(
                    f"data/{create_by.upper()}/Normal_Sign", order_id)
                if len(ls_file) > 0:
                    ls_file = ls_file.split(",")
                    if ls_file[-1] == "":
                        ls_file = ls_file[:-1]
                    else:
                        ls_file = ls_file
                    for file in ls_file:
                        wb = pd.read_excel(uploads_dir+"/"+file)
                        wb = wb.dropna(how='all')
                        wb = wb.fillna("")
                        wb = wb.reset_index(drop=True)
                        ls_table.append(wb.to_html(
                            index=False, classes='normal_sign tbl1 mb-0 table table-bordered'))

            else:
                uploads_dir = os.path.join(
                    f"data/{create_by.upper()}/Excel_Sign", order_id)
                ls_file = ls_file.split(",")
                wb = pd.read_excel(uploads_dir+"/"+ls_file[0])
                wb = wb.dropna(how='all')
                wb = wb.fillna("")
                wb = wb.reset_index(drop=True)
                wb.insert(0, 'Approve_all', '')
                ls_table.append(wb.to_html(
                    index=False, table_id="excel_sign", classes='excel_sign tbl1 mb-0 table table-bordered'))
            if order_id[:2] == "SO" or order_id[:2] == "De":
                if check_signer(order_id, mail):
                   return render_template('so_de_sign.html', order_id=order_id, context=context, ls_file=ls_file, tables=ls_table, file_style=file_style)
                else:
                    return render_template('signed.html')
            return render_template('esign.html', order_id=order_id, context=context, ls_file=ls_file, tables=ls_table, file_style=file_style)
        if order_id[:2] == "SO" or order_id[:2] == "De":
            if check_signer(order_id, mail):
                return render_template('so_de_sign.html', order_id=order_id, context=context, ls_file=ls_file, tables=ls_table, file_style=file_style)
            else:
                return render_template('signed.html')
        return render_template('esign.html', order_id=order_id, context=context)

@app.route('/', methods=['GET', 'POST'])
def login():
    session["type"] = ''

    # create the SSO oauth client
    oauthV2 = oauth.create_client('VN AI System')
    redirect_uri = url_for('authorize', _external=True)
    return oauthV2.authorize_redirect(redirect_uri)

# =================== login end redirect to SO system
@app.route('/So_login')
def so_login():
    for key in list(session.keys()):
        session.pop(key)
    session["type"] = 'SO'
    # create the SSO oauth client
    oauthV2 = oauth.create_client('VN AI System')
    redirect_uri = url_for('authorize', _external=True)
    return oauthV2.authorize_redirect(redirect_uri)

@app.route('/login')
def logins():
    for key in list(session.keys()):
        session.pop(key)
    session["type"] = 'login_out'
    # create the SSO oauth client
    oauthV2 = oauth.create_client('VN AI System')
    redirect_uri = url_for('authorize', _external=True)
    return oauthV2.authorize_redirect(redirect_uri)

@app.route('/login_out')
def login_out():
    data = session.get('datas')
    if data:
        return render_template('logins.html', data=data)
    else:
        return url_for(logins)

@app.route('/so')
def so():
    data = session.get('datas')
    if data:
        return render_template('so.html', data=data)
    else:
        return url_for('so_login')

# =================== login end redirect to Debit system
@app.route('/debit_login')
def debit_login():
    for key in list(session.keys()):
        session.pop(key)
    session["type"] = 'DEBIT'
    # create the SSO oauth client
    oauthV2 = oauth.create_client('VN AI System')
    redirect_uri = url_for('authorize', _external=True)
    return oauthV2.authorize_redirect(redirect_uri)

@app.route('/debit')
def debit():
    data = session.get('datas')
    if data:
        return render_template('debit.html', data=data)
    else:
        return url_for('debit_login')

@app.route('/authorize')
def authorize():
    data = []
    # create the SSO oauth client
    oauthV2 = oauth.create_client('VN AI System')
    # Access token from SSO (needed to get user info)
    token = oauthV2.authorize_access_token()
    userProfile = oauthV2.userinfo()  # uses openid endpoint to fetch user info
    # Here you use the profile/user data that you got and query your database find/register the user
    user_id = userProfile['username']
    name = userProfile['name']
    email = userProfile['email']

    if email != None:
        datas = {'username': userProfile['username'],
                 'name': userProfile['name'], 'email': userProfile['email']}
        data.append(user_id)
        data.append(name)
        data.append(email)
        if get_ID_mail(email) != user_id:
            create_user(str(data)[1:-1])
        # and set ur own data in the session not the profile from google
        session['username'] = userProfile['username']
        session['name'] = userProfile['name']
        # make the session permanant so it keeps existing after broweser gets closed
        session.permanent = True
        if not session.get("type"):
            if session.get("order_id") and session.get("url") == "esign" or session.get("url") == "":
                order_id = session.get('order_id')
                if order_id[:3] != "MPM":
                    return redirect(url_for('esign', order_id=order_id))
                else:
                    return redirect(url_for('mpm_esign', order_id=order_id))
            elif session.get("url") == "new_preview":
                return redirect(url_for('new_preview', order_id=session.get("order_id")[:-1], task=session.get("order_id")[-1]))
            elif session.get("url") == None:
                return redirect(url_for('newsign'))
            else:
                return redirect(url_for(f'{session.get("url")}'))
        else:
            if session.get("type") == "SO":
                session['datas'] = datas
                return redirect(url_for('so'))
            elif session.get("type") == "DEBIT":
                session['datas'] = datas
                return redirect(url_for('debit'))
            else:
                session['datas'] = datas
                return redirect(url_for('login_out'))
    else:
        return render_template('missing_mail.html')

@app.route('/comming')
def comming():
    if not session.get("username"):
        session["url"] = 'comming'
        return redirect(url_for('login'))
    else:
        return render_template('comming_soon.html')

@app.route('/contacts')
def contact():
    if not session.get("username"):
        session["url"] = 'contact'
        return redirect(url_for('login'))
    else:
        return render_template('comming_soon.html')

@app.route('/upfile/<order_id>', methods=['GET', 'POST'])
def upfile(order_id):
    ls_signer = []
    ls_file = ""
    file_style = ""

    if not session.get("username"):
        session["url"] = 'upfile'
        return redirect(url_for('login'))
    else:
        order_status = get_order('Order_Status', 'Order_No', order_id)[0]
        username = session.get("username")
        title = get_order('Order_Title', 'Order_No', order_id)[0]
        content = get_order('Description', 'Order_No', order_id)[0]
        for i in range(get_order('Signer_Qty', 'Order_No', order_id)[0]):
            ls_signer.append(
                get_order(f'Signer_ID{i+1}', 'Order_No', order_id)[0])
        context = {
            'order_id': order_id, 'ename': get_Name_ID(username), 'title': title, 'content': content
        }

        if request.method == 'POST':
            today = datetime.now()
            time = today.strftime("%m/%d/%Y %H:%M:%S")
            if len(request.files.getlist('files_nor')) >= 1:
                uploads_dir = os.path.join(
                    f"data/{username.upper()}/Normal_Sign", order_id)
                os.makedirs(uploads_dir, exist_ok=True)
                objects = os.listdir(uploads_dir)
                files_file = [f for f in objects if os.path.isfile(
                    os.path.join(uploads_dir, f))]
                for f in files_file:
                    os.remove(os.path.join(uploads_dir, f))
                for file in request.files.getlist('files_nor'):
                    if file.filename != "":
                        file.save(os.path.join(uploads_dir, file.filename))

                        ls_file += file.filename+","
                        file_style = 'excel'

                res = f"UPDATE Signature_Transaction SET File_name = '{ls_file}' , Updated_Time = '{time}', Updated_By = '{username}', File_Type = '{file_style}'  WHERE Order_No = '{order_id}'"
                update_order(res)
            elif request.files['file_exc'].filename != "":

                uploads_dir = os.path.join(
                    f"data/{username.upper()}/Excel_Sign", order_id)
                os.makedirs(uploads_dir, exist_ok=True)
                objects = os.listdir(uploads_dir)
                files_file = [f for f in objects if os.path.isfile(
                    os.path.join(uploads_dir, f))]
                for f in files_file:
                    os.remove(os.path.join(uploads_dir, f))
                file = request.files['file_exc']
                file.save(os.path.join(uploads_dir, file.filename))
                ls_file += file.filename + ","
                file_style = 'Sign_excel'

                res = f"UPDATE Signature_Transaction SET File_name = '{ls_file}' , Updated_Time = '{time}', Updated_By = '{username}',File_Type = '{file_style}'  WHERE Order_No = '{order_id}'"
                update_order(res)
            else:
                ls_file = ""
                res = f"UPDATE Signature_Transaction SET File_name = '{ls_file}' , Updated_Time = '{time}', Updated_By = '{username}',File_Type = '{file_style}'  WHERE Order_No = '{order_id}'"
                update_order(res)

            return redirect(url_for('preview', order_id=order_id))
    return render_template('upfile.html', ls_signer=ls_signer, context=context, order_status=order_status)

@app.route('/preview/<order_id>', methods=['GET', 'POST'])
@async_action
async def preview(order_id):
    ls_signer = []
    ls_time = []
    ls_status = []
    ls_table = []
    username = get_order('Create_By', 'Order_No', order_id)[0]
    ls_file = get_order('File_name', 'Order_No', order_id)[0]
    file_style = get_order('File_Type', 'Order_No', order_id)[0]
    session["order_id"] = order_id
    session["url"] = "esign"
    if get_order('Order_Status', 'Order_No', order_id)[0] == 'Reject':
        status = 'Rejected'
    elif get_order('Order_Status', 'Order_No', order_id)[0] == 'True':
        status = 'Signed'
    else:
        status = 'Waiting'
    for i in range(get_order('Signer_Qty', 'Order_No', order_id)[0]):
        ls_signer.append(get_order(f'Signer_ID{i+1}', 'Order_No', order_id)[0])
        ls_time.append(get_order(f'Time_Signer{i+1}', 'Order_No', order_id)[0])
        ls_status.append(
            get_order(f'Status_Sign{i+1}', 'Order_No', order_id)[0])
    context = {
        'employee': username, 'order_id': order_id, 'name': get_Name_ID(username), 'title': get_order(f'Order_Title', 'Order_No', order_id)[0], 'content': get_order(f'Description', 'Order_No', order_id)[0], 'status': status
    }

    if request.method == 'POST':
        if 'task' in request.form:
            session["order_id"] = None
            task = request.form['task']
            name = get_Name_mail(ls_signer[0])
            create_by =get_create_by(order_id)[0]
            if task == 'fn_prevSend':
                send_mail(ls_signer[0], get_order('Order_Title', 'Order_No', order_id)[
                          0], f'Dear User!<br>  You have 1 Order waiting to be signed!<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)}<br> Please click to this URL to Sign order http://10.220.40.75:5000/esign/{order_id}')
                return 'Done'
            else:
                await asyncio.sleep(600)
                send_mail(ls_signer[0], get_order('Order_Title', 'Order_No', order_id)[
                          0], f'Dear User!<br>  You have 1 Order waiting to be signed! <br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)}<br>  Please click to this URL to Sign order http://10.220.40.75:5000/esign/{order_id}')
                return 'Done'

    if ls_file != "":
        if file_style != "Sign_excel":
            uploads_dir = os.path.join(
                f"data/{username.upper()}/Normal_Sign", order_id)

            if len(ls_file) > 0:
                ls_file = ls_file.split(",")
                if ls_file[-1] == "":
                    ls_file = ls_file[:-1]
                else:
                    ls_file = ls_file
                for file in ls_file:
                    wb = pd.read_excel(uploads_dir+"/"+file)

                    wb = wb.dropna(how='all')
                    wb = wb.fillna("")
                    wb = wb.reset_index(drop=True)

                    ls_table.append(wb.to_html(
                        index=False, classes='normal_sign tbl1 mb-0 table table-bordered'))
        else:
            uploads_dir = os.path.join(
                f"data/{username.upper()}/Excel_Sign", order_id)
            ls_file = ls_file.split(",")
            wb = pd.read_excel(uploads_dir+"/"+ls_file[0])
            wb = wb.dropna(how='all')
            wb = wb.fillna("")
            wb = wb.reset_index(drop=True)
            ls_table.append(wb.to_html(index=False, classes='excel_sign tbl1 mb-0 table table-bordered'))
        return render_template('preview.html', ls_signer=ls_signer, context=context, ls_file=ls_file, ls_time=ls_time, ls_status=ls_status, tables=ls_table)
    else:
        return render_template('preview.html', ls_signer=ls_signer, context=context, ls_time=ls_time, ls_status=ls_status)

@app.route('/new_sign', methods=['GET', 'POST'])
def newsign():
    order_id = ""
    # sign_qty=""
    # order_title=""
    # descrip=""
    # file_link=""
    # file_style = ""
    # list_mail = []
    # list_signer = []
    files_name = []
    username = session.get("username")
    if not session.get("username"):
        session["url"] = 'newsign'
        return redirect(url_for('login'))
    else:
        session["url"] = 'newsign'
        if session.get("order_id"):
            if session.get("order_id")[-1] == "p":
                a = session.get("order_id")[:-1]
                delete = f"DELETE FROM Signature_Transaction WHERE Order_No = '{a}'"
                update_order(delete)
                mail = get_mail(username)
                context = {
                    'order_id': '', 'ename': get_Name_ID(username), 'mail': mail, 'username': username
                }
            else:
                mail = get_mail(username)
                context = {
                    'order_id': '', 'ename': get_Name_ID(username), 'mail': mail, 'username': username
                }
        else:
            mail = get_mail(username)
            context = {
                'order_id': order_id, 'ename': get_Name_ID(username), 'mail': mail, 'username': username
            }
    if request.method == 'POST':
        today = datetime.now()
        list_signer = []
        data = []
        time = today.strftime("%m/%d/%Y %H:%M:%S")
        order_id = 'CO'+str(today.strftime("%Y%m%d%H%M%S")
                            ) + str(random.randint(1, 999))
        ls_files = request.files.getlist('files')
        upload_dir = os.path.join(f"COST_UP_Data/{username}/{order_id}")
        os.makedirs(upload_dir, exist_ok=True)
        objects = os.listdir(upload_dir)
        files_file = [f for f in objects if os.path.isfile(
            os.path.join(upload_dir, f))]
        for f in files_file:
            os.remove(os.path.join(upload_dir, f))
        for file in ls_files:
            files_name.append(file.filename)
            file.save(os.path.join(upload_dir, file.filename))
        sl = 0
        for stt in range(10):
            if stt+1 <= 5:
                ls_mail = ""
                if request.form.getlist(f'email[{stt}]')[0] != "":
                    sl += 1
                for mail in request.form.getlist(f'email[{stt}]'):
                    if mail != "":
                        ls_mail += mail+","

                list_signer.append(ls_mail)
                list_signer.append("")
                list_signer.append("")
            else:
                list_signer.append("")
                list_signer.append("")
                list_signer.append("")

        data.append(order_id)
        data.append(username)
        data.append(time)
        data.append(f'Signed: 0/{sl}')
        data.append(sl)
        data += list_signer
        data.append(username)
        data.append(time)
        data.append(request.form['Title16Content'])
        data.append(request.form['Title17Content'])
        data.append('')
        data.append('excel')
        data.append(f'{files_name}')
        for a in range(32):
            data.append("")
        create_order(str(data)[1:-1])
        res = f"UPDATE Signature_Transaction SET Updated_Time = '{time}', Updated_By = '{username}', File_Type = 'excel'  WHERE Order_No = '{order_id}'"
        update_order(res)
        return redirect(url_for('new_preview', order_id=order_id, task="p"))

    return render_template('new_esign.html', context=context)

@app.route('/new_esign/<order_id><task>', methods=['GET', 'POST'])
def new_preview(order_id, task):
    ls_signer = []
    ls_time = []
    ls_status = []
    ls_name = []
    reas = []

    reas = get_order('Reason', 'Order_no', order_id)
    ls_style = ["Applicant", "Class supervisor", "Ministerial Director",
                "Audited by management", "Approved by supervisor", "BU Head"]
    username = session.get("username")
    file_name = get_order('File_name', 'Order_no', order_id)
    create_id = get_order('Create_By', 'Order_no', order_id)[0]

    creator = get_value('Name', 'User_Id', create_id)

    files_name = file_name[0][1:-1].split(",")
    title = get_order('Order_Title', 'Order_no', order_id)[0]
    content = get_order('Description', 'Order_no', order_id)[0]
    for i in range(get_order('Signer_Qty', 'Order_No', order_id)[0]):
        ls_signer.append(get_order(f'Signer_ID{i+1}', 'Order_No', order_id)[0])
        ls_time.append(get_order(f'Time_Signer{i+1}', 'Order_No', order_id)[0])
        ls_status.append(
            get_order(f'Status_Sign{i+1}', 'Order_No', order_id)[0])
    create_time = get_order('Create_Time', 'Order_no', order_id)[0]
    if not session.get("username"):
        session["order_id"] = order_id+task
        session["url"] = 'new_preview'
        return redirect(url_for('login'))
    else:
        session["url"] = 'new_preview'
        mail = get_mail(username)
        if check_signer(order_id, mail):
            for a in range(10):
                temp = sql_ress(
                    f"SELECT Status_Sign{a+1} FROM Signature_Transaction WHERE Signer_ID{a+1} LIKE '%{mail}%' AND Order_No = '{order_id}'")
                if temp != None:
                    if temp[0] == "" and (a+1) > 1:
                        temp2 = get_order(
                            f'Status_Sign{a}', 'Order_No', order_id)[0]
                        if temp2 != "":
                            loc = a+1
                        else:
                            loc = a
                        session["loc"] = loc
                    elif temp[0] == "" and (a+1) == 1:
                        session["loc"] = 1
        else:
            return render_template('new_signed.html')
        context = {
            'order_id': order_id, 'ename': get_Name_ID(username), 'mail': mail, 'username': username, 'title': title, 'content': content, 'creator': creator
        }
    for sign in ls_signer:
        if sign[:-1].find(',') == -1 and sign != "":
            ls_name.append(get_value('Name', 'Email', sign[:-1]))
        elif sign[:-1].find(',') != -1 and sign != "":
            na = ""
            for ml in sign[:-1].split(','):
                na += get_value('Name', 'Email', ml)+","
            ls_name.append(na[:-1])
    if task == "p":
        session["order_id"] = order_id+"p"
        if request.method == 'POST':
            create_by = get_create_by(order_id)[0]
            send_mail(ls_signer[0], get_order('Order_Title', 'Order_No', order_id)[
                      0], f'Dear User!<br>  You have 1 Order waiting to be signed!<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)}<br> Please click to this URL to Sign order http://10.220.40.75:5000/new_esign/{order_id}s')
            session.pop("order_id")
            return redirect(url_for("newsign"))
        return render_template('new_preview.html', context=context, files_name=files_name, ls_name=ls_name, ls_status=ls_status, ls_time=ls_time, create_time=create_time, task="preview", ls_style=ls_style)
    elif task == "v":
        session["order_id"] = order_id+"v"
        return render_template('new_preview.html', context=context, files_name=files_name, ls_name=ls_name, ls_status=ls_status, ls_time=ls_time, create_time=create_time, task="view", reas=reas, ls_style=ls_style)
    else:
        session["order_id"] = order_id+"s"
        if request.method == 'POST':
            lc = session.get("loc")
            reason = request.form['reason']
            if request.form['submit_button'] == "Reject":
                sign_ok(username, mail, order_id, "reject", lc, reason, reason)
                session.pop("order_id")
                return redirect(url_for("newsign"))
            else:
                sign_ok(username, mail, order_id,
                        "approve", lc, reason, reason)
                session.pop("order_id")
                return redirect(url_for("newsign"))

            # send_mail(ls_signer[0],get_order('Order_Title', 'Order_No', order_id)[0],f'Dear User!<br>  You have 1 Order waiting to be signed!<br> Please click to this URL to Sign order http://10.220.40.75:5000/esign/{order_id}')
            # return redirect(url_for("main"))
        return render_template('new_preview.html', context=context, files_name=files_name, ls_name=ls_name, ls_status=ls_status, ls_time=ls_time, create_time=create_time, task="sign", ls_style=ls_style)

@app.route('/view/<order_id>')
def view_app(order_id):
    ls_signer = []
    ls_time = []
    ls_status = []
    ls_name = []
    reas = []

    reas = get_order('Reason', 'Order_no', order_id)
    ls_style = ["申請人 / NGƯỜI XIN ĐƠN", "課級主管/Chủ quản cấp phòng", "部級主管/Chủ quản cấp bộ phận",
                "經管審核/Cost xét duyệt", "經管主管核准/Chủ quản Cost phê duyệt", "BU Head"]
    file_name = get_order('File_name', 'Order_no', order_id)
    create_id = get_order('Create_By', 'Order_no', order_id)[0]
    creator = get_value('Name', 'User_Id', create_id)
    mail = get_mail(create_id)
    files_name = file_name[0][1:-1].split(",")
    title = get_order('Order_Title', 'Order_no', order_id)[0]
    content = get_order('Description', 'Order_no', order_id)[0]
    for i in range(get_order('Signer_Qty', 'Order_No', order_id)[0]):
        ls_signer.append(get_order(f'Signer_ID{i+1}', 'Order_No', order_id)[0])
        ls_time.append(get_order(f'Time_Signer{i+1}', 'Order_No', order_id)[0])
        ls_status.append(
            get_order(f'Status_Sign{i+1}', 'Order_No', order_id)[0])
    create_time = get_order('Create_Time', 'Order_no', order_id)[0]
    context = {
        'order_id': order_id, 'ename': creator, 'mail': mail, 'username': create_id, 'title': title, 'content': content, 'creator': creator
    }
    for sign in ls_signer:
        if sign[:-1].find(',') == -1 and sign != "":
            ls_name.append(get_value('Name', 'Email', sign[:-1]))
        elif sign[:-1].find(',') != -1 and sign != "":
            na = ""
            for ml in sign[:-1].split(','):
                na += get_value('Name', 'Email', ml)+","
            ls_name.append(na[:-1])
    return render_template('new_preview.html', context=context, files_name=files_name, ls_name=ls_name, ls_status=ls_status, ls_time=ls_time, create_time=create_time, task="view", reas=reas, ls_style=ls_style)

@app.route('/new_main')
def new_main():
    ls_createby = []
    ls_createtime = []
    ls_title = []
    if not session.get("username"):
        session["url"] = 'new_main'
        return redirect(url_for('login'))
    else:
        session["url"] = 'new_main'
        username = session.get("username")
        mail = get_mail(username)
        ls_order = get_ls_co(mail)
        if ls_order:
            for id in ls_order:
                ls_createby.append(get_order('Create_By', 'Order_No', id)[0])
                ls_createtime.append(get_order('Create_Time', 'Order_No', id)[0])
                ls_title.append(get_order('Order_Title', 'Order_No', id)[0])
    return render_template('new_main.html', ls_order=ls_order, ls_createby=ls_createby, ls_createtime=ls_createtime, ls_title=ls_title)

@app.route('/co_get_all', methods= ['GET','POST'])
def new_search():
    ls_stt = []
    ls_Sign = []
    n_sign = []
    ls_createtime = []
    ls_signer = []
    ls_title = []
    new_ls = []
    ls_updatetime = []
    session["url"] = 'new_main'
    if not session.get("username"):

        return redirect(url_for('login'))
    else:

        username = session.get("username")

        list_order = sql_ress(
            f"SELECT Order_No FROM Signature_Transaction WHERE Create_By like '%{username}%' AND Order_No like 'CO%' ORDER BY Updated_Time DESC")
        if request.method == 'POST':
            task = request.form['task']
            alert_id = request.form['order_id']
            mail = get_Signer_now(alert_id)
            ls_title = get_order('Order_Title', 'Order_No', alert_id)[0]
            if task == 'alert':
                create_by = get_create_by(alert_id)[0]
                send_mail(mail, "(Alert!!!) " + ls_title,
                          f'Dear User!<br>  You have 1 Order waiting to be signed!<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)} <br> Please click to this URL to Sign order http://10.220.40.75:5000/new_esign/{alert_id}s')
                return 'done'
        if list_order != None:
            for i in list_order:
                new_ls.append(i)
                status = get_order('Order_Status', 'Order_No', i)[0]
                if status == "Reject":
                    ls_Sign.append(get_Signer_now(i))
                    ls_stt.append("Rejected!!!")
                elif status == "True":
                    ls_stt.append("Signed!!!")
                    ls_Sign.append(get_Signer_now(i))
                else:
                    ls_stt.append(f"{status}")
                    ls_Sign.append(get_Signer_now(i))

                ls_updatetime.append(
                    get_order('Updated_Time', 'Order_No', i)[0])
                ls_createtime.append(
                    get_order('Create_Time', 'Order_No', i)[0])
                ls_title.append(get_order('Order_Title', 'Order_No', i)[0])
            a = ""
            b = ""
            for sign in ls_Sign:
                if sign[:-1].find(",") == -1 and sign[:-1] != "":
                    n_sign.append(sign[:-1])
                    ls_signer.append(get_Name_mail(sign[:-1]))
                elif sign[:-1] != "":

                    for sm in sign[:-1].split(","):
                        if sm != "":
                            if get_Name_mail(sm) != None:
                                a += get_Name_mail(sm) + ","
                                b += sm + ","
                            else:
                                a += "None,"
                                b += sm + ","
                    n_sign.append(b[:-1])
                    ls_signer.append(a[:-1])
            return render_template('new_search.html', ls_updatetime=ls_updatetime, list_order=new_ls, ls_stt=ls_stt, ls_Sign=n_sign, ls_createtime=ls_createtime, ls_signer=ls_signer, ls_title=ls_title,task = 'all')
        else:
            return render_template('new_search.html')

@app.route('/co_get_waiting', methods= ['GET','POST'])
def new_get_waiting():
    ls_stt = []
    ls_Sign = []
    n_sign = []
    ls_createtime = []
    ls_signer = []
    ls_title = []
    new_ls = []
    ls_updatetime = []
    session["url"] = 'new_get_waiting'
    if not session.get("username"):

        return redirect(url_for('login'))
    else:

        username = session.get("username")

        list_order = sql_ress(
            f"SELECT Order_No FROM Signature_Transaction WHERE Create_By like '%{username}%' AND Order_No like 'CO%' ORDER BY Updated_Time DESC")
        if request.method == 'POST':
            task = request.form['task']
            alert_id = request.form['order_id']
            mail = get_Signer_now(alert_id)
            ls_title = get_order('Order_Title', 'Order_No', alert_id)[0]
            if task == 'alert':
                create_by = get_create_by(alert_id)[0]
                send_mail(mail, "(Alert!!!) " + ls_title,
                          f'Dear User!<br>  You have 1 Order waiting to be signed!<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)} <br> Please click to this URL to Sign order http://10.220.40.75:5000/new_esign/{alert_id}s')
                return 'done'
        if list_order != None:
            for i in list_order:
                new_ls.append(i)
                status = get_order('Order_Status', 'Order_No', i)[0]
                if status == "Reject":
                    ls_Sign.append(get_Signer_now(i))
                    ls_stt.append("Rejected!!!")
                elif status == "True":
                    ls_stt.append("Signed!!!")
                    ls_Sign.append(get_Signer_now(i))
                else:
                    ls_stt.append(f"{status}")
                    ls_Sign.append(get_Signer_now(i))

                ls_updatetime.append(
                    get_order('Updated_Time', 'Order_No', i)[0])
                ls_createtime.append(
                    get_order('Create_Time', 'Order_No', i)[0])
                ls_title.append(get_order('Order_Title', 'Order_No', i)[0])
            a = ""
            b = ""
            for sign in ls_Sign:
                if sign[:-1].find(",") == -1 and sign[:-1] != "":
                    n_sign.append(sign[:-1])
                    ls_signer.append(get_Name_mail(sign[:-1]))
                elif sign[:-1] != "":

                    for sm in sign[:-1].split(","):
                        if sm != "":
                            if get_Name_mail(sm) != None:
                                a += get_Name_mail(sm) + ","
                                b += sm + ","
                            else:
                                a += "None,"
                                b += sm + ","
                    n_sign.append(b[:-1])
                    ls_signer.append(a[:-1])
            return render_template('new_search.html', ls_updatetime=ls_updatetime, list_order=new_ls, ls_stt=ls_stt, ls_Sign=n_sign, ls_createtime=ls_createtime, ls_signer=ls_signer, ls_title=ls_title,task = 'waiting')
        else:
            return render_template('new_search.html')

@app.route('/co_get_signed', methods= ['GET','POST'])
def new_get_signed():
    ls_stt = []
    ls_Sign = []
    n_sign = []
    ls_createtime = []
    ls_signer = []
    ls_title = []
    new_ls = []
    ls_updatetime = []
    session["url"] = 'new_get_signed'
    if not session.get("username"):

        return redirect(url_for('login'))
    else:

        username = session.get("username")

        list_order = sql_ress(
            f"SELECT Order_No FROM Signature_Transaction WHERE Create_By like '%{username}%' AND Order_No like 'CO%' ORDER BY Updated_Time DESC")
        if request.method == 'POST':
            task = request.form['task']
            alert_id = request.form['order_id']
            mail = get_Signer_now(alert_id)
            ls_title = get_order('Order_Title', 'Order_No', alert_id)[0]
            if task == 'alert':
                create_by = get_create_by(alert_id)[0]
                send_mail(mail, "(Alert!!!) " + ls_title,
                          f'Dear User!<br>  You have 1 Order waiting to be signed!<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)} <br> Please click to this URL to Sign order http://10.220.40.75:5000/new_esign/{alert_id}s')
                return 'done'
        if list_order != None:
            for i in list_order:
                new_ls.append(i)
                status = get_order('Order_Status', 'Order_No', i)[0]
                if status == "Reject":
                    ls_Sign.append(get_Signer_now(i))
                    ls_stt.append("Rejected!!!")
                elif status == "True":
                    ls_stt.append("Signed!!!")
                    ls_Sign.append(get_Signer_now(i))
                else:
                    ls_stt.append(f"{status}")
                    ls_Sign.append(get_Signer_now(i))

                ls_updatetime.append(
                    get_order('Updated_Time', 'Order_No', i)[0])
                ls_createtime.append(
                    get_order('Create_Time', 'Order_No', i)[0])
                ls_title.append(get_order('Order_Title', 'Order_No', i)[0])
            a = ""
            b = ""
            for sign in ls_Sign:
                if sign[:-1].find(",") == -1 and sign[:-1] != "":
                    n_sign.append(sign[:-1])
                    ls_signer.append(get_Name_mail(sign[:-1]))
                elif sign[:-1] != "":

                    for sm in sign[:-1].split(","):
                        if sm != "":
                            if get_Name_mail(sm) != None:
                                a += get_Name_mail(sm) + ","
                                b += sm + ","
                            else:
                                a += "None,"
                                b += sm + ","
                    n_sign.append(b[:-1])
                    ls_signer.append(a[:-1])
            return render_template('new_search.html', ls_updatetime=ls_updatetime, list_order=new_ls, ls_stt=ls_stt, ls_Sign=n_sign, ls_createtime=ls_createtime, ls_signer=ls_signer, ls_title=ls_title,task = 'signed')
        else:
            return render_template('new_search.html')

@app.route('/co_get_rejected')
def new_get_rejected():
    ls_stt = []
    ls_Sign = []
    n_sign = []
    ls_createtime = []
    ls_signer = []
    ls_title = []
    new_ls = []
    ls_updatetime = []
    session["url"] = 'new_get_rejected'
    if not session.get("username"):

        return redirect(url_for('login'))
    else:

        username = session.get("username")

        list_order = sql_ress(
            f"SELECT Order_No FROM Signature_Transaction WHERE Create_By like '%{username}%' AND Order_No like 'CO%' ORDER BY Updated_Time DESC")
        if request.method == 'POST':
            task = request.form['task']
            alert_id = request.form['order_id']
            mail = get_Signer_now(alert_id)
            ls_title = get_order('Order_Title', 'Order_No', alert_id)[0]
            if task == 'alert':
                create_by = get_create_by(alert_id)[0]
                send_mail(mail, "(Alert!!!) " + ls_title,
                          f'Dear User!<br>  You have 1 Order waiting to be signed!<br> Order create by:{create_by} -- {get_Name_ID(create_by)}<br>Email: {get_mail(create_by)}<br>Phone: {get_phone(create_by)} <br> Please click to this URL to Sign order http://10.220.40.75:5000/new_esign/{alert_id}s')
                return 'done'
        if list_order != None:
            for i in list_order:
                new_ls.append(i)
                status = get_order('Order_Status', 'Order_No', i)[0]
                if status == "Reject":
                    ls_Sign.append(get_Signer_now(i))
                    ls_stt.append("Rejected!!!")
                elif status == "True":
                    ls_stt.append("Signed!!!")
                    ls_Sign.append(get_Signer_now(i))
                else:
                    ls_stt.append(f"{status}")
                    ls_Sign.append(get_Signer_now(i))

                ls_updatetime.append(
                    get_order('Updated_Time', 'Order_No', i)[0])
                ls_createtime.append(
                    get_order('Create_Time', 'Order_No', i)[0])
                ls_title.append(get_order('Order_Title', 'Order_No', i)[0])
            a = ""
            b = ""
            for sign in ls_Sign:
                if sign[:-1].find(",") == -1 and sign[:-1] != "":
                    n_sign.append(sign[:-1])
                    ls_signer.append(get_Name_mail(sign[:-1]))
                elif sign[:-1] != "":

                    for sm in sign[:-1].split(","):
                        if sm != "":
                            if get_Name_mail(sm) != None:
                                a += get_Name_mail(sm) + ","
                                b += sm + ","
                            else:
                                a += "None,"
                                b += sm + ","
                    n_sign.append(b[:-1])
                    ls_signer.append(a[:-1])
            return render_template('new_search.html', ls_updatetime=ls_updatetime, list_order=new_ls, ls_stt=ls_stt, ls_Sign=n_sign, ls_createtime=ls_createtime, ls_signer=ls_signer, ls_title=ls_title,task = 'reject')
        else:
            return render_template('new_search.html')

@app.route('/show_file/<file_names>')
def show_file(file_names):
    order_id = file_names.split(':')[0]
    user_id = get_order('Create_By', 'Order_no', order_id)[0]
    Path_file = os.path.join(f"COST_UP_Data/{user_id}/{order_id}")
    file_name = file_names.split(':')[1]
    file_type = file_name.split(".")[-1]
    if file_type == "pdf":
        return send_file(os.path.join(Path_file, file_name), as_attachment=False)
    else:
        try:
            df = pd.read_excel(os.path.join(Path_file, file_name))
        except:
            df = pd.read_csv(os.path.join(Path_file, file_name))
        return df.to_html(classes='normal_sign mb-0 table table-bordered', index=False)

@app.route('/download/<file_name>')
def download(file_name):
    order_id = file_name.split(':')[0]
    user_id = get_order('Create_By', 'Order_no', order_id)[0]
    Path_file = os.path.join(f"COST_UP_Data/{user_id}/{order_id}")
    return send_file(Path_file+"/"+file_name.split(':')[1], as_attachment=True)

@app.route('/no_sign', methods=['GET', 'POST'])
def no_sign():
    order_id = ""
    sign_qty = ""
    order_title = ""
    descrip = ""
    file_link = ""
    file_style = ""
    list_mail = []
    list_signer = []
    data = []

    today = datetime.now()
    time = today.strftime("%m/%d/%Y %H:%M:%S")
    username = session.get("username")
    print(request.environ['REMOTE_ADDR'] + " == " + time)
    if not session.get("username"):
        return redirect(url_for('login'))
    else:
        if session.get("order_id"):
            a = session.get("order_id")
            delete = f"DELETE FROM Signature_Transaction WHERE Order_No = '{a}'"
            update_order(delete)
            mail = get_mail(username)

            context = {
                'order_id': '', 'ename': get_Name_ID(username), 'mail': mail
            }
        else:
            mail = get_mail(username)
            context = {
                'order_id': order_id, 'ename': get_Name_ID(username), 'mail': mail
            }
        if request.method == 'POST':
            today = datetime.now()
            time = today.strftime("%m/%d/%Y %H:%M:%S")
            order_id = str(today.strftime("%Y%m%d%H%M%S")) + \
                str(random.randint(1, 999))
            ls_mail = request.form.getlist('ls_mail')
            order_title = request.form['title']
            descrip = request.form['content']
            ls_file = ""
            sign_qty = len(set(ls_mail))-1
            for mail in sorted(set(ls_mail), key=ls_mail.index):
                if mail != "":
                    list_mail.append(mail)
            if len(request.files.getlist('files_nor')) >= 1:
                uploads_dir = os.path.join(
                    f"data/{username.upper()}/Normal_Sign", order_id)
                os.makedirs(uploads_dir, exist_ok=True)
                objects = os.listdir(uploads_dir)
                files_file = [f for f in objects if os.path.isfile(
                    os.path.join(uploads_dir, f))]
                for f in files_file:
                    os.remove(os.path.join(uploads_dir, f))
                for file in request.files.getlist('files_nor'):
                    if file.filename != "":
                        file.save(os.path.join(uploads_dir, file.filename))

                        ls_file += file.filename+","
                        file_style = 'excel'
                    else:
                        file_style = 'excel'
            for stt in range(10):
                if stt+1 <= sign_qty:
                    list_signer.append(list_mail[stt])
                    list_signer.append("")
                    list_signer.append("")
                else:
                    list_signer.append("")
                    list_signer.append("")
                    list_signer.append("")

            data.append(order_id)
            data.append(username)
            data.append(time)
            data.append(f'Signed: 0/{sign_qty}')
            data.append(sign_qty)
            data += list_signer
            data.append(username)
            data.append(time)
            data.append(order_title)
            data.append(descrip)
            data.append(file_link)
            data.append(file_style)
            for a in range(33):
                data.append("")
            create_order(str(data)[1:-1])
            res = f"UPDATE Signature_Transaction SET File_name = '{ls_file}' , Updated_Time = '{time}', Updated_By = '{username}', File_Type = '{file_style}'  WHERE Order_No = '{order_id}'"
            update_order(res)
            return redirect(url_for('preview', order_id=order_id))
        return render_template('no_sign.html', context=context)

@app.route('/excel_sign', methods=['GET', 'POST'])
def excel_sign():
    order_id = ""
    sign_qty = ""
    order_title = ""
    descrip = ""
    file_link = ""
    file_style = ""
    list_mail = []
    list_signer = []
    data = []

    username = session.get("username")
    print(request.environ['REMOTE_ADDR'])
    if not session.get("username"):
        return redirect(url_for('login'))
    else:
        if session.get("order_id"):
            a = session.get("order_id")
            delete = f"DELETE FROM Signature_Transaction WHERE Order_No = '{a}'"
            update_order(delete)
            mail = get_mail(username)

            context = {
                'order_id': '', 'ename': get_Name_ID(username), 'mail': mail
            }
        else:
            mail = get_mail(username)
            context = {
                'order_id': order_id, 'ename': get_Name_ID(username), 'mail': mail
            }
        if request.method == 'POST':
            today = datetime.now()
            time = today.strftime("%m/%d/%Y %H:%M:%S")
            order_id = str(today.strftime("%Y%m%d%H%M%S")) + \
                str(random.randint(1, 999))
            ls_mail = request.form.getlist('ls_mail')
            order_title = request.form['title']
            descrip = request.form['content']
            ls_file = ""
            sign_qty = len(set(ls_mail))-1
            for mail in sorted(set(ls_mail), key=ls_mail.index):
                if mail != "":
                    list_mail.append(mail)
            if request.files['file_exc'].filename != "":
                uploads_dir = os.path.join(
                    f"data/{username.upper()}/Excel_Sign", order_id)
                os.makedirs(uploads_dir, exist_ok=True)
                objects = os.listdir(uploads_dir)
                files_file = [f for f in objects if os.path.isfile(
                    os.path.join(uploads_dir, f))]
                for f in files_file:
                    os.remove(os.path.join(uploads_dir, f))
                file = request.files['file_exc']
                file.save(os.path.join(uploads_dir, file.filename))
                ls_file += file.filename
                file_style = 'Sign_excel'
            for stt in range(10):
                if stt+1 <= sign_qty:
                    list_signer.append(list_mail[stt])
                    list_signer.append("")
                    list_signer.append("")
                else:
                    list_signer.append("")
                    list_signer.append("")
                    list_signer.append("")

            data.append(order_id)
            data.append(username)
            data.append(time)
            data.append(f'Signed: 0/{sign_qty}')
            data.append(sign_qty)
            data += list_signer
            data.append(username)
            data.append(time)
            data.append(order_title)
            data.append(descrip)
            data.append(file_link)
            data.append(file_style)
            for a in range(33):
                data.append("")
            create_order(str(data)[1:-1])
            res = f"UPDATE Signature_Transaction SET File_name = '{ls_file}' , Updated_Time = '{time}', Updated_By = '{username}', File_Type = '{file_style}'  WHERE Order_No = '{order_id}'"
            update_order(res)
            return redirect(url_for('preview', order_id=order_id))
        return render_template('excel_sign.html', context=context)

@app.route('/mail', methods=['GET', 'POST'])
def index():
    order_id = ""
    sign_qty = ""
    order_title = ""
    descrip = ""
    file_link = ""
    file_style = ""
    list_mail = []
    list_signer = []
    data = []

    username = session.get("username")
    print(request.environ['REMOTE_ADDR'])
    if not session.get("username"):
        return redirect(url_for('login'))
    else:
        if session.get("order_id"):
            a = session.get("order_id")
            delete = f"DELETE FROM Signature_Transaction WHERE Order_No = '{a}'"
            update_order(delete)
            mail = get_mail(username)

            context = {
                'order_id': '', 'ename': get_Name_ID(username), 'mail': mail
            }
        else:
            mail = get_mail(username)
            context = {
                'order_id': order_id, 'ename': get_Name_ID(username), 'mail': mail
            }

    if request.method == 'POST':
        today = datetime.now()
        time = today.strftime("%m/%d/%Y %H:%M:%S")
        order_id = str(today.strftime("%Y%m%d%H%M%S")) + \
            str(random.randint(1, 999))
        ls_mail = request.form.getlist('ls_mail')
        order_title = request.form['title']
        descrip = request.form['content']
        ls_file = ""
        sign_qty = len(set(ls_mail))-1
        for mail in sorted(set(ls_mail), key=ls_mail.index):
            if mail != "":
                list_mail.append(mail)
        if request.files['file_exc'].filename != "":
            uploads_dir = os.path.join(
                f"data/{username.upper()}/Excel_Sign", order_id)
            os.makedirs(uploads_dir, exist_ok=True)
            objects = os.listdir(uploads_dir)
            files_file = [f for f in objects if os.path.isfile(
                os.path.join(uploads_dir, f))]
            for f in files_file:
                os.remove(os.path.join(uploads_dir, f))
            file = request.files['file_exc']
            file.save(os.path.join(uploads_dir, file.filename))
            ls_file += file.filename
            file_style = 'Sign_excel'
        elif len(request.files.getlist('files_nor')) >= 1:
            uploads_dir = os.path.join(
                f"data/{username.upper()}/Normal_Sign", order_id)
            os.makedirs(uploads_dir, exist_ok=True)
            objects = os.listdir(uploads_dir)
            files_file = [f for f in objects if os.path.isfile(
                os.path.join(uploads_dir, f))]
            for f in files_file:
                os.remove(os.path.join(uploads_dir, f))
            for file in request.files.getlist('files_nor'):
                if file.filename != "":
                    file.save(os.path.join(uploads_dir, file.filename))

                    ls_file += file.filename+","
                    file_style = 'excel'
                else:
                    file_style = 'excel'
        for stt in range(10):
            if stt+1 <= sign_qty:
                list_signer.append(list_mail[stt])
                list_signer.append("")
                list_signer.append("")
            else:
                list_signer.append("")
                list_signer.append("")
                list_signer.append("")

        data.append(order_id)
        data.append(username)
        data.append(time)
        data.append(f'Signed: 0/{sign_qty}')
        data.append(sign_qty)
        data += list_signer
        data.append(username)
        data.append(time)
        data.append(order_title)
        data.append(descrip)
        data.append(file_link)
        data.append(file_style)
        for a in range(33):
            data.append("")
        create_order(str(data)[1:-1])
        res = f"UPDATE Signature_Transaction SET File_name = '{ls_file}' , Updated_Time = '{time}', Updated_By = '{username}', File_Type = '{file_style}'  WHERE Order_No = '{order_id}'"
        update_order(res)
        return redirect(url_for('preview', order_id=order_id))
    return render_template('index.html', context=context)

@app.route('/so_sign', methods = ['GET','POST'])
def so_sign():
    ls_createby_ex = []
    ls_createtime_ex = []
    ls_title_ex = []
    new_ls = []
    if not session.get("username"):
        session["url"] = "so_sign"
        return redirect(url_for('login'))
    else:
        session["url"] = "so_sign"
        username = session.get("username")
        mail = get_mail(username)
        # ls_order = get_ls_order_sign(mail)
        ls_ex_order = get_ls_ex(mail)

        for i in ls_ex_order:
            if i[:2] == 'SO':
                new_ls.append(i)
                ls_createby_ex.append(get_order('Create_By', 'Order_No', i)[0])
                ls_createtime_ex.append(get_order('Create_Time', 'Order_No', i)[0])
                ls_title_ex.append(get_order('Order_Title', 'Order_No', i)[0])
        # for id in ls_order:
        #     ls_createby.append(get_order('Create_By', 'Order_No', id)[0])
        #     ls_createtime.append(get_order('Create_Time', 'Order_No', id)[0])
        #     ls_title.append(get_order('Order_Title', 'Order_No', id)[0])
    return render_template('main.html',   ls_order=new_ls, ls_createby=ls_createby_ex, ls_createtime=ls_createtime_ex, ls_title=ls_title_ex)

@app.route('/debit_sign', methods = ['GET','POST'])
def debit_sign():
    ls_createby_ex = []
    ls_createtime_ex = []
    ls_title_ex = []
    new_ls = []
    if not session.get("username"):
        session["url"] = "debit_sign"
        return redirect(url_for('login'))
    else:
        session["url"] = "debit_sign"
        username = session.get("username")
        mail = get_mail(username)
        # ls_order = get_ls_order_sign(mail)
        ls_ex_order = get_ls_ex(mail)

        for i in ls_ex_order:
            if i[:2] == 'De':
                new_ls.append(i)
                ls_createby_ex.append(get_order('Create_By', 'Order_No', i)[0])
                ls_createtime_ex.append(get_order('Create_Time', 'Order_No', i)[0])
                ls_title_ex.append(get_order('Order_Title', 'Order_No', i)[0])
        # for id in ls_order:
        #     ls_createby.append(get_order('Create_By', 'Order_No', id)[0])
        #     ls_createtime.append(get_order('Create_Time', 'Order_No', id)[0])
        #     ls_title.append(get_order('Order_Title', 'Order_No', id)[0])
    return render_template('main.html',   ls_order=new_ls, ls_createby=ls_createby_ex, ls_createtime=ls_createtime_ex, ls_title=ls_title_ex)

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if not session.get("username"):
        return redirect(url_for('login'))
    else:
        path_name = session.get("username")
        uploads_dir = os.path.join(app.instance_path, str(path_name))
        os.makedirs(uploads_dir, exist_ok=True)
        if request.method == 'POST':
            # save the single "profile" file
            profile = request.files['file']
            if profile.filename != "":
                objects = os.listdir(uploads_dir)
                if len(objects) > 0:
                    files_file = [f for f in objects if os.path.isfile(
                        os.path.join(uploads_dir, f))]
                    os.remove(os.path.join(uploads_dir, files_file[0]))
                profile.save(os.path.join(uploads_dir, profile.filename))
                return redirect(url_for('upload'))
            return render_template('upload.html')
        else:
            objects = os.listdir(uploads_dir)
            if len(objects) > 0:
                files_file = [f for f in objects if os.path.isfile(
                    os.path.join(uploads_dir, f))]
                df = pd.read_excel(os.path.join(uploads_dir, files_file[0]))
                message = ""
                try:
                    # df = df[['P/N','Project','usage','Prod stage']]
                    df = df[['P/N', 'Prod stage', 'Project', 'usage']]
                    df['usage'] = pd.to_numeric(
                        round(df['usage']), downcast='integer')

                except Exception as err:
                    message = err
                # if message == "":
                #     pn_list= np.unique(df['P/N'].values)
                #     for pn_id in pn_list:

                #         list_p =  df[ df['P/N'].isin([pn_id]) ]
                #     # list_p["usage"] = pd.to_numeric(list_p["usage"] , downcast='integer')
                #     list_p = list_p.astype({"usage": int})

                #     list_p = list_p[['P/N', 'Prod stage','Project','usage' ]]

                return render_template('upload.html', tables=[df.to_html(classes='normal_sign mb-0 table table-bordered', header=False, index=False, table_id='tbl')])
        return render_template('upload.html')

@app.route('/home')
def home():
    return render_template('home.html')

@app.route('/logout')
def logout():
    if session.get("url") == "newsign" or session.get("url") == "new_main":
        for key in list(session.keys()):
            session.pop(key)
        return redirect(url_for('newsign'))
    else:
        for key in list(session.keys()):
            session.pop(key)
        return redirect('/home')


# ===============================END SIGNATURE=================================

if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True, port=5000)
