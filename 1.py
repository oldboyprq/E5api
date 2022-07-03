# -*- coding: UTF-8 -*-
import requests as req
import json,sys,time
#先注册azure应用,确保应用有以下权限:[应用程序委托]
#files:	Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
#user:	User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
#mail:  Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
#注册后一定要再点代表xxx授予管理员同意,否则outlook api无法调用






path=sys.path[0]+r'/token.txt'

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
        print("reading....")
        r = req.get(r'https://graph.microsoft.com/v1.0/me/drive/sharedWithMe', headers=headers)
        if r.status_code == 200:
            print("1、与我共享的文件调用成功")
            value = json.loads(r.text)
            print("   共有%d封"%len(value['value']))
        r = req.get(r'https://graph.microsoft.com/v1.0/me/', headers=headers)
        if r.status_code == 200:
            print("2、我的个人资料调用成功")
        r = req.get(r'https://graph.microsoft.com/v1.0/me/messages?top=1000', headers=headers)
        if r.status_code == 200:
            print("3、我的邮件调用成功")
            value = json.loads(r.text)
            print("   共有%d封" % len(value['value']))
        r = req.get(r'https://graph.microsoft.com/v1.0/me/insights/trending', headers=headers)
        if r.status_code == 200:
            print("4、我的常用项目调用成功")
        r = req.get(r'https://graph.microsoft.com/v1.0/me/calendars', headers=headers)
        if r.status_code == 200:
            print("5、我的日历调用成功")
        r = req.get(r'https://graph.microsoft.com/v1.0/me/messages?$search="hello world"',
                    headers=headers)
        if r.status_code == 200:
            print("6、我的包含helloworld的邮件调用成功")
            value = json.loads(r.text)
            print("   共有%d封" % len(value['value']))
        r = req.get(r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories', headers=headers)
        if r.status_code == 200:
            print("7、我的outlook类别调用成功")
        r = req.get(r'https://graph.microsoft.com/v1.0/me/drive/recent', headers=headers)
        if r.status_code == 200:
            print("8、我最近使用的文件调用成功")
            value = json.loads(r.text)
            print("   共有%d个" % len(value['value']))
        r = req.get(r'https://graph.microsoft.com/v1.0/me/people', headers=headers)
        if r.status_code == 200:
            print("9、与我合作的人员调用成功")
            value = json.loads(r.text)
            print("   共有%d人" % len(value['value']))
        print("writing....")
         # 创建文件夹
        data = {
            "name": "New Folder",
            "folder": {},
        }
        r = req.post(r"https://graph.microsoft.com/v1.0/me/drive/root/children", data=json.dumps(data), headers=headers)
        if r.status_code < 300:
            print("1、新建文件夹成功")
        files_id = json.loads(r.text)['id']
        # 删除文件夹
        r = req.delete(r"https://graph.microsoft.com/v1.0/me/drive/items/%s" % files_id, headers=headers)
        if r.status_code < 300:
            print("2、删除文件夹成功")
        # 更新txt
        r = req.put(r"https://graph.microsoft.com/v1.0/me/drive/root:/api/api.txt:/content", headers=headers, data="api files")
        if r.status_code < 300:
            print("3、txt内容更新成功")       
        #发送邮件
        mailmessage = {
                        "message": {
                            "subject": "Meet for lunch? hello world",
                            "body": {
                                "contentType": "Text",
                                "content": "The new cafeteria is open.hello world."
                            },
                            "toRecipients": [
                                {
                                    "emailAddress": {
                                        "address": "lx@rosepig.onmicrosoft.com"
                                    }
                                },
                                {
                                    "emailAddress": {
                                        "address": "lzj@rosepig.onmicrosoft.com"
                                    }
                                },
                                {
                                    "emailAddress": {
                                        "address": "pig@rosepig.onmicrosoft.com"
                                    }
                                }
                            ]
                        }
                    }
        # 只发一次邮件
        r = req.post(r'https://graph.microsoft.com/v1.0/me/sendMail', headers=headers, data=json.dumps(mailmessage))
        if r.status_code == 202:
            print("1、测试邮件发送成功")
        print("此轮运行时间为{}-{}-{} {}:{}:{}".format(x[0], x[1], x[2], x[3] + 8, x[4], x[5]))
    except Exception as e:
        print("something error")
        print(e)


main()
