# -*- coding: UTF-8 -*-
import requests as req
import json,sys,time
#先注册azure应用,确保应用有以下权限:
#files:	Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
#user:	User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
#mail:  Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
#注册后一定要再点代表xxx授予管理员同意,否则outlook api无法调用






path=sys.path[0]+r'/1.txt'

def gettoken(refresh_token):
    headers={'Content-Type':'application/x-www-form-urlencoded'
            }
    data={'grant_type': 'refresh_token',
          'refresh_token': refresh_token,
          'client_id':id,
          'client_secret':secret,
          'redirect_uri':'http://localhost:53682/'
         }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    with open(path, 'w+') as f:
        f.write(refresh_token)
    return access_token
def main():
    fo = open(path, "r+")
    refresh_token = fo.read()
    fo.close()
    x=time.localtime(time.time())
    access_token=gettoken(refresh_token)
    headers={
    'Authorization':access_token,
    'Content-Type':'application/json'
    }
    try:
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/sharedWithMe',headers=headers).status_code == 200:
            print("1、与我共享的文件调用成功")
        if req.get(r'https://graph.microsoft.com/v1.0/me/',headers=headers).status_code == 200:
            print("2、我的个人资料调用成功")
        if req.get(r'https://graph.microsoft.com/v1.0/me/messages',headers=headers).status_code == 200:
            print("3、我的邮件调用成功")
        if req.get(r'https://graph.microsoft.com/v1.0/me/insights/trending',headers=headers).status_code == 200:
            print("4、我的常用项目调用成功")
        if req.get(r'https://graph.microsoft.com/v1.0/me/calendars',headers=headers).status_code == 200:
            print("5、我的日历调用成功")
        if req.get(r'https://graph.microsoft.com/v1.0/me/messages?$search="hello world"',headers=headers).status_code == 200:
            print("6、我的包含helloworld的邮件调用成功")
        if req.get(r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories',headers=headers).status_code == 200:
            print("7、我的outlook类别调用成功")
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/recent',headers=headers).status_code == 200:
            print("8、我最近使用的文件调用成功")
        if req.get(r'https://graph.microsoft.com/v1.0/me/people',headers=headers).status_code == 200:
            print("9、与我合作的人员调用成功")
        print("此次运行时间为{}-{}-{} {}:{}:{}".format(x[0],x[1],x[2],x[3]+8,x[4],x[5]))
    except:
        print("pass")
        pass
for _ in range(2):
    main()
