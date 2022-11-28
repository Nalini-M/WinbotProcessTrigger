from imap_tools import MailBox, AND
import requests
import json
import pandas as pd
import cx_Oracle
import keyring
from exchangelib import DELEGATE, Account,Identity, Credentials,Version,Configuration,OAuth2Credentials,OAUTH2
from exchangelib.autodiscover import clear_cache
clear_cache()
def getAuthentication(granttype, clientid, clientsecret, clientscope, OrchestratorURL):
    print("inside get auth")
    url = "" + OrchestratorURL + "/identity/connect/token"
    data = "grant_type=client_credentials&client_id="+clientid+"&client_secret="+clientsecret+"&scope="+clientscope
    # url = "https://winfo101.winfosolutions.com/api/Account/Authenticate"

    # data = {"tenancyName": "" + tenantName + "", "usernameOrEmailAddress": "" + username + "","password": "" + password + ""}

    header = {"content-type": "application/x-www-form-urlencoded"}

    response = requests.post(url, data=data, headers=header, verify=False)
    print(str(response.status_code))

    r_json = response.json()

    key = r_json["access_token"]
    print("key is " + str(key))
    return key


def getRobotId(key, inputRobotEnv, folderid, OrchestratorUrl):
    print("in get robotid orche url-->" + OrchestratorUrl)

    url = "" + OrchestratorUrl + "/odata/Robots"
    print(url)
    print(key)

    header = {"content-type": "application/json",
              "Authorization": "Bearer " + str(key),
              "X-UIPATH-OrganizationUnitId": ""+folderid+""}    #folder id is 2 for test env
    response = requests.get(url, headers=header, verify=False)
    print(str(response.status_code))

    r_json = response.json()
    print(str(r_json))
    val = json.dumps(r_json['value'])
    print("here")
    print(val)
    resp_dict = json.loads(val)
    for i in resp_dict:
        print("in for-->" + inputRobotEnv)
        if inputRobotEnv in i['RobotEnvironments']:

            print("in if")

            robotId = i['Id']
            robotName = str(i['Name'])
            print(
                "now let us check whether status is available or not for this following robot-->" + robotName + " " + str(
                    robotId))
            robot_status = getRobotStatus(key, robotId, folderid, OrchestratorUrl)
            print("stats-->" + robot_status)
            if robot_status == 'Available':
                print("this robot is available-->" + robotName)
                return robotId, robot_status
            else:
                print("this robot is busy -->" + robotName)
            print("Robot id-->" + str(robotId))
            print(robotName)
    return "noid", "nothing"


def getRobotStatus(key, RobotID, folderid, OrchestratorUrl):
    global Statee
    print("in get robot status function")
    url = "" + OrchestratorUrl + "/odata/Sessions?$filter=Robot/Id eq " + str(RobotID) + "&$ Select =State"

    header = {"content-type": "application/json", "Authorization": "Bearer " + str(key),
              "X-UIPATH-OrganizationUnitId": ""+folderid+""}
    response = requests.get(url, headers=header, verify=False)

    print(str(response.status_code))

    r_json = response.json()
    print(str(r_json))
    val = json.dumps(r_json['value'])
    print("here")
    print(val)
    resp_dict = json.loads(val)
    for i in resp_dict:
        Statee = i['State']
        # print(Status)
    return Statee


def getReleaseKey(key, inputProcessKey, folderid, OrchestratorUrl):
    global releaseKey
    url = "" + OrchestratorUrl + "/odata/Releases"

    header = {"content-type": "application/json",
              "Authorization": "Bearer " + str(key),
              "X-UIPATH-OrganizationUnitId": ""+folderid+""}

    response = requests.get(url, headers=header, verify=False)
    print(str(response.status_code))

    r_json = response.json()
    val = json.dumps(r_json['value'])
    resp_dict = json.loads(val)
    print(resp_dict)
    for i in resp_dict:
        if i['ProcessKey'] == inputProcessKey:
            releaseKey = str(i['Key'])
            processKey = str(i['ProcessKey'])
            print("processKey is " + str(processKey))
            print("releaseKey is " + str(releaseKey))
    return releaseKey

    # startJob(key, robotId, releaseKey)


def startJob(key, robotId, releaseKey, OrchestratorUrl, folderid,Trigger_Point,Process_Name,Instance_Name):
    print("key is " + str(key))
    print("robot id is " + str(robotId))
    print("release key is " + str(releaseKey))

    url = "" + OrchestratorUrl + "/odata/Jobs/UiPath.Server.Configuration.OData.StartJobs"

    data = {
        "startInfo": {
            "ReleaseKey": str(releaseKey),
            "Strategy": "Specific",
            "RobotIds": [robotId],
            "InputArguments": "{\"Trigger_Point\":\""+Trigger_Point+"\",\"Process_Name\":\""+Process_Name+"\",\"Application_Name\":\""+Instance_Name+"\",\"RPA_ID\":\"""\",\"Retry\":\"No\",\"FOLDER_PATH\":\"""\"}"
            #"InputArguments": credentials
        }
    }

    header = {"content-type": "application/json",
              "Authorization": "Bearer " + str(key),
              "X-UIPATH-OrganizationUnitId": ""+folderid+""}

    response = requests.post(url, data=json.dumps(data), headers=header, verify=False)
    print(str(response.status_code))

    r_json = response.json()
    jobId = r_json['value'][0]['Id']
    print(str(jobId))


def PWDKEYRING(NetworkAddress, user_name):
    return keyring.get_password(NetworkAddress, user_name)


def MailCheck():
    with open('C:\\WINBOTFlaskApplication\\ProcessTrigger\\Config.txt') as config:
        json_data = json.load(config)
        # ------------------------------Only database details will be in the Config file
        databasePort = json_data["Port"]
        databaseHostName = json_data["HostName"]
        DatabaseServiceName = json_data["ServiceName"]
        DatabaseUsername = json_data["User"]
        DBNetworkAddress = json_data["DBNetworkAddress"]
        Databasepassword = PWDKEYRING(DBNetworkAddress, DatabaseUsername)
        try:
            dsn = cx_Oracle.makedsn(
                databaseHostName,
                databasePort,
                service_name=DatabaseServiceName
            )
            print('initiated connection to oracle')
            conn = cx_Oracle.connect(
                user=DatabaseUsername,
                password=Databasepassword,
                dsn=dsn
            )
            c = conn.cursor()
            #-------------------insert new tables queries here
            sql_query = 'SELECT * FROM WB_CONFIG_GROUP'
            print(sql_query)
            c.execute(sql_query)
            print('sql query executed')
            #fetch header column names
            headercolumns = [x[0] for x in c.description]  # for getting column names
            print(headercolumns)
            #fetch all rows
            headerrows = c.fetchall()
            if not headerrows:
                print('no rows')
            else:
                headerdata = pd.DataFrame(headerrows, columns=headercolumns)  # it will give output as same as table
                print(headerdata)
                for header_row in headerdata.itertuples():
                    if header_row.CONFIG_GROUP_NAME == 'UiPath Orchestrator':
                        sql_query = 'SELECT * FROM WB_CONFIG_LINES where config_group_id='+str(header_row.CONFIG_GROUP_ID)
                        print(sql_query)
                        c.execute(sql_query)
                        print('sql query executed')
                        # fetch header column names
                        headercolumnslines = [x[0] for x in c.description]  # for getting column names
                        print(headercolumnslines)
                        # fetch all rows
                        headerlinesrows = c.fetchall()
                        headerlinedata = pd.DataFrame(headerlinesrows, columns=headercolumnslines)
                        granttype = ''
                        clientid = headerlinedata['CONFIGURATION_VALUE'][headerlinedata['CONFIGURATION_NAME'] == 'Orchestrator Client Details'].values[0]
                        clientsecret = headerlinedata['CONFIG_PASSWORD'][headerlinedata['CONFIGURATION_NAME'] == 'Orchestrator Client Details'].values[0]
                        clientscope = headerlinedata['CONFIGURATION_VALUE'][headerlinedata['CONFIGURATION_NAME'] == 'Orchestrator Scope'].values[0]
                        OrchestratorURL = headerlinedata['CONFIGURATION_VALUE'][headerlinedata['CONFIGURATION_NAME'] == 'Orchestrator URL'].values[0]
                        folderid = headerlinedata['CONFIGURATION_VALUE'][headerlinedata['CONFIGURATION_NAME'] == 'Orchestrator Folder Id'].values[0]
                        print("check")
                        print(clientid,clientsecret,clientscope,OrchestratorURL,folderid)
                        print("done")
                    elif header_row.CONFIG_GROUP_NAME == 'Email':
                        sql_query = 'SELECT * FROM WB_CONFIG_LINES where config_group_id=' +str(header_row.CONFIG_GROUP_ID)
                        print(sql_query)
                        c.execute(sql_query)
                        print('sql query executed')
                        # fetch header column names
                        headercolumnslines = [x[0] for x in c.description]  # for getting column names
                        print(headercolumnslines)
                        # fetch all rows
                        headerlinesrows = c.fetchall()
                        headerlinedata = pd.DataFrame(headerlinesrows, columns=headercolumnslines)
                        mail = headerlinedata['CONFIGURATION_VALUE'][
                            headerlinedata['CONFIGURATION_NAME'] == 'Inbound Email Address'].values[0]
                        emailaddress = headerlinedata['CONFIGURATION_VALUE'][
                            headerlinedata['CONFIGURATION_NAME'] == 'Email Network Address'].values[0]
                        servername = headerlinedata['CONFIGURATION_VALUE'][
                            headerlinedata['CONFIGURATION_NAME'] == 'Server Name'].values[0]
                        imapport = headerlinedata['CONFIGURATION_VALUE'][
                            headerlinedata['CONFIGURATION_NAME'] == 'Imap Port'].values[0]
                    elif header_row.CONFIG_GROUP_NAME == 'Sharepoint':
                        sql_query = 'SELECT * FROM WB_CONFIG_LINES where config_group_id=' +str(header_row.CONFIG_GROUP_ID)
                        print(sql_query)
                        c.execute(sql_query)
                        print('sql query executed')
                        # fetch header column names
                        headercolumnslines = [x[0] for x in c.description]  # for getting column names
                        print(headercolumnslines)
                        # fetch all rows
                        headerlinesrows = c.fetchall()
                        headerlinedata = pd.DataFrame(headerlinesrows, columns=headercolumnslines)
                        sharepointclientsecret = headerlinedata['CONFIG_PASSWORD'][
                            headerlinedata['CONFIGURATION_NAME'] == 'Sharepoint Client Details'].values[0]
                        sharepointclientid = headerlinedata['CONFIGURATION_VALUE'][
                            headerlinedata['CONFIGURATION_NAME'] == 'Sharepoint Client Details'].values[0]
                        sharepointusername = headerlinedata['CONFIGURATION_VALUE'][
                            headerlinedata['CONFIGURATION_NAME'] == 'Sharepoint Login Credentials'].values[0]
                        sharepointpassword = headerlinedata['CONFIG_PASSWORD'][
                            headerlinedata['CONFIGURATION_NAME'] == 'Sharepoint Login Credentials'].values[0]
                        sharepointscope = headerlinedata['CONFIGURATION_VALUE'][
                            headerlinedata['CONFIGURATION_NAME'] == 'Sharepoint Scope'].values[0]
                        sharepointtenantname = headerlinedata['CONFIGURATION_VALUE'][
                            headerlinedata['CONFIGURATION_NAME'] == 'Sharepoint Tenant'].values[0]


                pwd = PWDKEYRING(emailaddress, mail)

                sql_query = 'SELECT * FROM PROCESS_ADMINISTRATION where enable = \'Yes\''
                print(sql_query)
                c.execute(sql_query)
                # fetch header column names
                linecolumns = [x[0] for x in c.description]  # for getting column names
                # fetch all rows
                linerows = c.fetchall()
                if not linerows:
                    print('no rows')
                else:
                    linedata = pd.DataFrame(linerows,
                                            columns=linecolumns)
                #---------------------------------------------------
                for i in linedata.itertuples():
                    if i.TRG_SOURCE == 'Sharepoint':
                        url_token = "https://login.microsoftonline.com/" + str(
                            sharepointtenantname) + "/oauth2/v2.0/token"
                        print("before data token ")
                        # data_token = "grant_type=client_credentials&username=winbot@winfosolutions.compassword=Even@odd&client_id=1f024f5b-1169-4d2c-80dc-67f199cd6d20&client_secret=nZk8Q~LMuU1g-bZBXUVz2jSl5Fc1gq5uP_3nGbfO&scope=https://graph.microsoft.com/.default"

                        print("spu", sharepointusername)
                        print("pswd", sharepointpassword)
                        print("client id ", sharepointclientid)
                        print("client secret ", sharepointclientsecret)
                        print("scope ", sharepointscope)
                        inputRobotEnv = i.ORCH_TRG_ENV_NAME
                        data_token = "grant_type=client_credentials&username="+str(sharepointusername)+"&password="+str(sharepointpassword)+"&client_id="+str(sharepointclientid)+"&client_secret="+str(sharepointclientsecret)+"&scope="+str(sharepointscope)
                        print("data token ", data_token)
                        header_token = {"content-type": "application/x-www-form-urlencoded",
                                        "SdkVersion": "postman-graph/v1.0"}

                        response_token = requests.post(url_token, data=data_token, headers=header_token, verify=False)
                        print('code')
                        print(str(response_token.status_code))
                        AccessToken_json = response_token.json()
                        Sharepoint_AccessToken = AccessToken_json["access_token"]
                        site_name = 'winbot'
                        print("Sharepoint access token is " + str(Sharepoint_AccessToken))
                        url_site = "https://graph.microsoft.com/v1.0/sites/winfoconsulting.sharepoint.com:/sites/" + site_name
                        header_site = {"Authorization": "Bearer " + str(Sharepoint_AccessToken)}
                        response_site = requests.get(url_site, headers=header_site, verify=False)
                        s_json = response_site.json()
                        print(s_json)
                        site_id = s_json['id']

                        print("site_id is " + site_id)
                        url_drive_id = "https://graph.microsoft.com/v1.0/sites/" + site_id + "/drives"
                        print("url drive id ", url_drive_id)
                        header_drive_id = {"Authorization": "Bearer " + str(Sharepoint_AccessToken)}
                        response_drive_id = requests.get(url_drive_id, headers=header_drive_id, verify=False)
                        d_json = response_drive_id.json()
                        print("drive id response ", d_json)
                        val = json.dumps(d_json['value'])
                        print(val)
                        resp_dict = json.loads(val)
                        print(str(resp_dict[0]['id']))
                        drive_id = resp_dict[0]['id']
                        print("drive_id is ", drive_id)
                        folder_value = i.DOC_REP_INPUT_FOLDER+"/"+"Inbox"
                        folder_value_arr = folder_value.split("/")
                        move_folder_value = i.DOC_REP_INPUT_FOLDER+"/"+"Input"
                        # inputProcessKey = i.ORCH_PROCESS_TRG_NAME
                        # inputRobotEnv = i.ORCH_TRG_ENV_NAME
                        print("move folder value " + str(move_folder_value))
                        move_folder_value_arr = move_folder_value.split("/")
                        count = 0
                        parent_folder_id = ""
                        move_parent_folder_id = ""
                        files_present = False
                        for folder_name in folder_value_arr:
                            if count == 0:
                                url = "https://graph.microsoft.com/v1.0/drives/" + drive_id + "/root:/" + folder_name
                                print("URl in parent folder " + str(url))
                                header = {"Authorization": "Bearer " + str(Sharepoint_AccessToken)}
                                response = requests.get(url, headers=header, verify=False)
                                p_json = response.json()
                                parent_folder_id = p_json['id']
                            else:
                                url = "https://graph.microsoft.com/v1.0/drives/" + drive_id + "/items/" + parent_folder_id + "/children?filter=name eq '" + folder_name + "'"
                                print("URl in parent folder " + str(url))
                                header = {"Authorization": "Bearer " + str(Sharepoint_AccessToken)}
                                response = requests.get(url, headers=header, verify=False)
                                p_json = response.json()
                                val = json.dumps(p_json['value'])
                                print(val)
                                resp_dict_p = json.loads(val)
                                print(str(resp_dict_p[0]['id']))
                                parent_folder_id = resp_dict_p[0]['id']

                            # header = {"Authorization": "Bearer " + str(Sharepoint_AccessToken)}
                            # response = requests.get(url, headers=header, verify=False)
                            # p_json = response.json()
                            # print("folder respone ",r_json)
                            # print("folder id ",response.json()['id'])
                            # val = json.dumps(r_json['value'])
                            # print(val)
                            # resp_dict = json.loads(val)
                            # print(str(resp_dict[0]['id']))
                            # parent_folder_id = resp_dict[0]['id']

                            count = count + 1
                            print("parent folder id is " + str(parent_folder_id))
                            count_m = 0
                        url_input = "https://graph.microsoft.com/v1.0/drives/" + drive_id + "/items/" + parent_folder_id + "/children"
                        header_input = {"Authorization": "Bearer " + str(Sharepoint_AccessToken)}
                        response_input = requests.get(url_input, headers=header_input, verify=False)
                        r_i_json = response_input.json()
                        val_p = json.dumps(r_i_json['value'])
                        print("parent folder children response ", val_p)
                        resp_dict_p = json.loads(val_p)
                        count_files = 0
                        for x in resp_dict_p:
                            print("x value is ", x)
                            folder_id = x['id']
                            folder_name = x['name']
                            print("folder id ", folder_id)
                            print("folder name ", folder_name)
                            file_move_path_val = ""
                            file_move_path_val = move_folder_value
                            file_move_path_arr = file_move_path_val.split("/")
                            count = 0
                            for folder_name_val in file_move_path_arr:
                                if count == 0:
                                    url = "https://graph.microsoft.com/v1.0/drives/" + drive_id + "/root:/" + folder_name_val
                                    print("URl in move parent folder if " + str(url))
                                    header = {"Authorization": "Bearer " + str(Sharepoint_AccessToken)}
                                    response = requests.get(url, headers=header, verify=False)
                                    p_json = response.json()
                                    move_parent_folder_id = p_json['id']
                                else:
                                    url = "https://graph.microsoft.com/v1.0/drives/" + drive_id + "/items/" + move_parent_folder_id + "/children?filter=name eq '" + folder_name_val + "'"
                                    print("URl in move parent folder else " + str(url))
                                    header = {"Authorization": "Bearer " + str(Sharepoint_AccessToken)}
                                    response = requests.get(url, headers=header, verify=False)
                                    p_json = response.json()
                                    val = json.dumps(p_json['value'])
                                    print(val)
                                    resp_dict_p = json.loads(val)
                                    print(str(resp_dict_p[0]['id']))
                                    move_parent_folder_id = resp_dict_p[0]['id']
                                count = count + 1
                            url_fc = "https://graph.microsoft.com/v1.0/drives/" + drive_id + "/items/" + parent_folder_id + "/children"
                            header_fc = {"Authorization": "Bearer " + str(Sharepoint_AccessToken)}
                            response_fc = requests.get(url_fc, headers=header_fc, verify=False)
                            r_json_fc = response_fc.json()
                            print(r_json_fc)
                            val_fc = json.dumps(r_json_fc['value'])
                            resp_dict_fc = json.loads(val_fc)
                            file_val = "file"
                            for y in resp_dict_fc:
                                if file_val in y:
                                    count_files = count_files + 1
                                print("count files", count_files)
                                if count_files > 0:
                                    print("in count files", count_files)
                                    token = getAuthentication(granttype, clientid, clientsecret, clientscope,
                                                              OrchestratorURL)
                                    RobotStateAndID = getRobotId(token, inputRobotEnv, folderid,
                                                                 OrchestratorURL)

                                    files_present = True
                            if RobotStateAndID[1] == 'Available':
                                for y in resp_dict_fc:
                                    file_id = y['id']
                                    file_name = y['name']
                                    url3 = "https://graph.microsoft.com/v1.0/drives/" + drive_id + "/items/" + file_id
                                    print(url3)
                                    data3 = {
                                        """parentReference""": {
                                            """id""": "" + str(move_parent_folder_id) + ""},
                                        """name""": "" + str(file_name) + ""

                                    }

                                    # data3 = """{"parentReference": {"id":"01Q6PN2TTAFLKIAXYTZVELGPUW7ZJ7NBQZ"}, "name":"""""+file_name+"""""}"""
                                    print(json.dumps(data3))
                                    header3 = {"content-type": "application/json",
                                               "Authorization": "Bearer " + str(Sharepoint_AccessToken)}
                                    response3 = requests.patch(url3, headers=header3, data=json.dumps(data3),
                                                               verify=False)
                                    print(response3.status_code)
                                    print(response3.reason)
                                    print(file_id)
                                    inputProcessKey = i.ORCH_PROCESS_TRG_NAME

                        if files_present == True:
                            print("hmm")
                            token = getAuthentication(granttype, clientid, clientsecret, clientscope,
                                                      OrchestratorURL)
                            robotID = getRobotId(token, inputRobotEnv, folderid, OrchestratorURL)
                            releaseKey = getReleaseKey(token, inputProcessKey, folderid, OrchestratorURL)
                            startJob(token, robotID[0], releaseKey, OrchestratorURL, folderid, i.TRG_SOURCE,
                                     i.WB_PROCESS_NAME, i.TARGET_APPLICATION)







                    elif i.TRG_SOURCE == 'Mail':
                        print("mail->" + str(mail))
                        print("pwd ->" + str(pwd))
                        credentials = OAuth2Credentials(client_id='69a713a0-fe37-4bf2-bb20-39ca9ea7e178', client_secret='oKH8Q~ugA9jE~Teq0PxeSPZ8P2PTxuRMgMW-Ob7m', tenant_id='7c51d221-e93b-4397-b4c7-775faf9f6d10', identity=Identity(smtp_address='winbot@winfosolutions.com'))
                        config = Configuration(server='outlook.office365.com', credentials=credentials,  auth_type=OAUTH2)
                        account = Account('winbot@winfosolutions.com', access_type=DELEGATE, config=config)
                        #account = Account('winbot@winfosolutions.com', credentials=credentials, autodiscover=True)
                        #creds = OAuth2Credentials(
                        #client_id='69a713a0-fe37-4bf2-bb20-39ca9ea7e178' ,client_secret='oKH8Q~ugA9jE~Teq0PxeSPZ8P2PTxuRMgMW-Ob7m',  tenant_id='7c51d221-e93b-4397-b4c7-775faf9f6d10')
                        #credentials = Credentials(username='winbot@winfosolutions.com', password='Even@odd')
                        #config = Configuration(server='outlook.office365.com', credentials=credentials)
                        #account = Account(primary_smtp_address=mail, config=config,
                                    #autodiscover=False, access_type=DELEGATE)

                        #creds = Credentials(username=mail, password=pwd)
                        
                        #config = Configuration(credentials=creds, auth_type=OAUTH2,service_point='https://outlook.office365.com/owa/')
                        print('jj')
                        #account = Account(primary_smtp_address=mail,  autodiscover=False,config=config,access_type=DELEGATE)
                        print('ff')
                        unread_mail_list = account.inbox.filter(is_read=False)
                        #with MailBox(servername, imapport).login(mail, pwd, initial_folder='INBOX') as mailbox:
                        print("in mail box")
                        breakFor = False

                        for msg in unread_mail_list:
                            print("fetching the mails")
                            mailExists = False

                            print("all->" + msg.subject)

                            if not linerows:
                                print('no rows')
                            else:
                                linedata = pd.DataFrame(linerows,
                                                        columns=linecolumns)  # it will give output as same as table
                                for i in linedata.itertuples():
                                    if i.TRG_SOURCE == 'Mail':
                                        mailExists = False
                                        print("subj--->" + i.MAIL_SUBJECT)
                                        subj = i.MAIL_SUBJECT
                                        if subj in msg.subject:
                                            # inputRobotName = i.ROBOT_NAME
                                            inputProcessKey = i.ORCH_PROCESS_TRG_NAME
                                            inputRobotEnv = i.ORCH_TRG_ENV_NAME
                                            mailFolder = i.DOC_REP_INPUT_FOLDER
                                            token = getAuthentication(granttype, clientid, clientsecret,
                                                                      clientscope,
                                                                      OrchestratorURL)
                                            RobotStateAndID = getRobotId(token, inputRobotEnv, folderid,
                                                                         OrchestratorURL)
                                            # RobotState = getRobotStatus(token, robotID, OrchestratorURL)
                                            print(
                                                "RobotState after getting the return statement from authentication to main code -->" + str(
                                                    RobotStateAndID[1]))
                                            if RobotStateAndID[1] == 'Available':
                                                mailExists = True
                                                mailfolder1 = mailFolder
                                                # mailbox.move(msg.uid, mailfolder1)
                                                print(mailfolder1)
                                                # print(account.root.tree())
                                                to_folder = account.root / 'Top of Information Store' / mailfolder1
                                                for msg1 in unread_mail_list:
                                                    print('h123')
                                                    print(msg1.subject)
                                                    print(subj)
                                                    if subj in msg1.subject:
                                                        msg1.move(to_folder)

                                            else:
                                                print("Robot is busy for this subject-->" + str(subj))
                                                # breakFor=True
                                            break
                                if breakFor == True:
                                    break
                                if mailExists == True:
                                    token = getAuthentication(granttype, clientid, clientsecret, clientscope,
                                                              OrchestratorURL)
                                    robotID = getRobotId(token, inputRobotEnv, folderid, OrchestratorURL)
                                    releaseKey = getReleaseKey(token, inputProcessKey, folderid, OrchestratorURL)
                                    startJob(token, robotID[0], releaseKey, OrchestratorURL, folderid,i.TRG_SOURCE,i.WB_PROCESS_NAME,i.TARGET_APPLICATION)
                                    print("mails read")
                                    print("next turn")
                                    break



        except cx_Oracle.Error as error:
            print(error)
        finally:
            # release the connection
            if conn:
                conn.close()


if __name__ == '__main__':
    MailCheck()
