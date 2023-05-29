import requests
import imaplib

## 定义发件人地址和密码以及收件人地址信息
username = '<mailbox address>'

client_id='<Application (client) ID>'
client_secret='<client secret>'
tenant_id='<Directory (tenant) ID>'

# Get a token
url = 'https://login.partner.microsoftonline.cn/'+tenant_id+'/oauth2/v2.0/token'
data = {
    
    'grant_type': 'client_credentials',
    'client_id': client_id,
    'client_secret': client_secret,
    'scope': 'https://partner.outlook.cn/.default',
}
res = requests.post(url, data=data)
print("请求响应结果", res)
token = res.json().get('access_token')
print("访问令牌", token)

def generate_auth_string(user, token):
    return f"user={user}\x01auth=Bearer {token}\x01\x01"

# 连接IMAP服务器并获取邮件
try:
    imap_conn = imaplib.IMAP4_SSL('partner.outlook.cn', 993)
    imap_conn.debug = 4
    imap_conn.authenticate("XOAUTH2", lambda x:generate_auth_string(username,token))
    imap_conn.select('Inbox')
    print (imap_conn.list())

except imaplib.IMAP4.error as e:
    print("邮件发送失败", e)
