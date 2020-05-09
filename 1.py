# -*- coding: UTF-8 -*-
import requests as req
import json,sys,time
#先注册azure应用,确保应用有以下权限:
#files:	Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
#user:	User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
#mail:  Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
#注册后一定要再点代表xxx授予管理员同意,否则outlook api无法调用






path=sys.path[0]+r'/1.txt'
num1 = 0
num2 = 0
roundnum = 0
totalroundnum = 100
failnum = 0
totalfailnum = 0

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
    global num1
    global num2
    global roundnum
    global totalroundnum
    global localtime
    global failnum
    global totalfailnum
    num2 = 0
    failnum = 0
    localtime = time.asctime( time.localtime(time.time()) )
    access_token=gettoken(refresh_token)
    headers={
    'Authorization':access_token,
    'Content-Type':'application/json'
    }
    try:
        api1 = req.get(r'https://graph.microsoft.com/v1.0/me/drive/root',headers=headers).status_code
        if api1 == 200:
            num1+=1
            num2+=1
            print(':) Success ['+str(api1)+'] - graph.microsoft.com/v1.0/me/drive/root')
        else:
            print(':( Failure ['+str(api1)+'] - graph.microsoft.com/v1.0/me/drive/root')
            failnum += 1
            totalfailnum += 1
        api2 = req.get(r'https://graph.microsoft.com/v1.0/me/drive',headers=headers).status_code
        if api2 == 200:
            num1+=1
            num2+=1
            print(':) Success ['+str(api2)+'] - graph.microsoft.com/v1.0/me/drive')
        else:
            print(':( Failure ['+str(api2)+'] - graph.microsoft.com/v1.0/me/drive')
            failnum += 1
            totalfailnum += 1
        api3 = req.get(r'https://graph.microsoft.com/v1.0/drive/root',headers=headers).status_code
        if api3 == 200:
            num1+=1
            num2+=1
            print(':) Success ['+str(api3)+'] - graph.microsoft.com/v1.0/drive/root')
        else:
            print(':( Failure ['+str(api3)+'] - graph.microsoft.com/v1.0/drive/root')
            failnum += 1
            totalfailnum += 1
        api4 = req.get(r'https://graph.microsoft.com/v1.0/users ',headers=headers).status_code
        if api4 == 200:
            num1+=1
            num2+=1
            print(':) Success ['+str(api4)+'] - graph.microsoft.com/v1.0/users')
        else:
            print(':( Failure ['+str(api4)+'] - graph.microsoft.com/v1.0/users')
            failnum += 1
            totalfailnum += 1
        api5 = req.get(r'https://graph.microsoft.com/v1.0/me/messages',headers=headers).status_code
        if api5 == 200:
            num1+=1
            num2+=1
            print(':) Success ['+str(api5)+'] - graph.microsoft.com/v1.0/me/messages')  
        else:
            print(':( Failure ['+str(api5)+'] - graph.microsoft.com/v1.0/me/messages')  
            failnum += 1
            totalfailnum += 1
        api6 = req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',headers=headers).status_code
        if api6 == 200:
            num1+=1
            num2+=1
            print(':) Success ['+str(api6)+'] - graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules')   
        else:
            print(':( Failure ['+str(api6)+'] - graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules') 
            failnum += 1
            totalfailnum += 1
        api7 = req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox',headers=headers).status_code
        if api7 == 200:
            num1+=1
            num2+=1
            print(':) Success ['+str(api7)+'] - graph.microsoft.com/v1.0/me/mailFolders/inbox')
        else:
            print(':( Failure ['+str(api7)+'] - graph.microsoft.com/v1.0/me/mailFolders/inbox')
            failnum += 1
            totalfailnum += 1
        api8 = req.get(r'https://graph.microsoft.com/v1.0/me/drive/root/children',headers=headers).status_code
        if api8 == 200:
            num1+=1
            num2+=1
            print(':) Success ['+str(api8)+'] - graph.microsoft.com/v1.0/me/drive/root/children')
        else:
            print(':( Failure ['+str(api8)+'] - graph.microsoft.com/v1.0/me/drive/root/children')
            failnum += 1
            totalfailnum += 1
        api9 = req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders',headers=headers).status_code
        if api9 == 200:
            num1+=1
            num2+=1
            print(':) Success ['+str(api9)+'] - graph.microsoft.com/v1.0/me/mailFolders')
        else:
            print(':( Failure ['+str(api9)+'] - graph.microsoft.com/v1.0/me/mailFolders')
            failnum += 1
            totalfailnum += 1
        api10 = req.get(r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories',headers=headers).status_code
        if api10 == 200:
            num1+=1
            num2+=1
            print(':) Success ['+str(api10)+'] - graph.microsoft.com/v1.0/me/outlook/masterCategories')
        else:
            print(':( Failure ['+str(api10)+'] - graph.microsoft.com/v1.0/me/outlook/masterCategories')
            failnum += 1
            totalfailnum += 1
    except:
        print(':o Something went wrong.')
        pass
    else:
        print(':D Test completed.')
        print('-- Success: '+str(num2))
        print('-- Failure: '+str(failnum))
for _ in range(totalroundnum):
    roundnum += 1
    print('')
    print(str(roundnum)+' / '+str(totalroundnum))
    print('----------')
    main()
    print('----------')
    print('Total success: '+str(num1))
    print('Total failure: '+str(totalfailnum))
    print('Current time of server: ', localtime+'')
    print('')
