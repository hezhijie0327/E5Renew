import requests, json, sys

# Current Version: 1.0.0

## How to get and use?
# git clone "https://github.com/hezhijie0327/E5Renew.git" && python ./E5Renew/E5Renew.py

## Parameter
client_id = r"config_client_id"
client_secret = r"config_client_secret"
round_ID = 0
refresh_token_file = sys.path[0] + r"/refresh_token.txt"

## Function
# Get Token
def GetToken(refresh_token):
    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }
    data = {
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": "http://localhost:53682/"
    }
    html = requests.post("https://login.microsoftonline.com/common/oauth2/v2.0/token", data = data, headers = headers)
    response = json.loads(html.text)
    refresh_token = response["refresh_token"]
    access_token = response["access_token"]
    with open(refresh_token_file, "w+") as refresh_token_file_write:
        refresh_token_file_write.write(refresh_token)
    return access_token
# Main
def main():
    global round_ID
    round_ID = round_ID + 1
    refresh_token_file_read = open(refresh_token_file, "r+")
    refresh_token = refresh_token_file_read.read()
    refresh_token_file_read.close()
    access_token = GetToken(refresh_token)
    headers = {
        "Authorization": access_token,
        "Content-Type": "application/json"
    }
    try:
        print("{")
        print("    " + "\"round_ID\": " + str(round_ID) + ",")
        # Permission: Calendars.Read
        if requests.get(r"https://graph.microsoft.com/v1.0/me/calendars", headers = headers).status_code == 200:
            print("    " + "\"calendars_read\": " + "true" + ",")
        else:
            print("    " + "\"calendars_read\": " + "false" + ",")
        # Permission: Contacts.Read
        if requests.get(r"https://graph.microsoft.com/v1.0/me/contacts", headers = headers).status_code == 200:
            print("    " + "\"contacts_read\": " + "true" + ",")
        else:
            print("    " + "\"contacts_read\": " + "false" + ",")
        # Permission: Directory.Read.All
        if requests.get(r"https://graph.microsoft.com/v1.0/me/memberOf", headers = headers).status_code == 200:
            print("    " + "\"directory_read_all\": " + "true" + ",")
        else:
            print("    " + "\"directory_read_all\": " + "false" + ",")
        # Permission: Files.Read.All
        if requests.get(r"https://graph.microsoft.com/v1.0/drive", headers = headers).status_code == 200:
            print("    " + "\"files_read_all\": " + "true" + ",")
        else:
            print("    " + "\"files_read_all\": " + "false" + ",")
        # Permission: Group.Read.All
        if requests.get(r"https://graph.microsoft.com/v1.0/groups", headers = headers).status_code == 200:
            print("    " + "\"group_read_all\": " + "true" + ",")
        else:
            print("    " + "\"group_read_all\": " + "false" + ",")
        # Permission: Mail.Read
        if requests.get(r"https://graph.microsoft.com/v1.0/me/messages", headers = headers).status_code == 200:
            print("    " + "\"mail_read\": " + "true" + ",")
        else:
            print("    " + "\"mail_read\": " + "false" + ",")
        # Permission: Notes.Read.All
        if requests.get(r"https://graph.microsoft.com/v1.0/me/onenote/notebooks", headers = headers).status_code == 200:
            print("    " + "\"notes_read_all\": " + "true" + ",")
        else:
            print("    " + "\"notes_read_all\": " + "false" + ",")
        # Permission: People.Read.All
        if requests.get(r"https://graph.microsoft.com/v1.0/me/people", headers = headers).status_code == 200:
            print("    " + "\"people_read_all\": " + "true" + ",")
        else:
            print("    " + "\"people_read_all\": " + "false" + ",")
        # Permission: Sites.Read.All
        if requests.get(r"https://graph.microsoft.com/v1.0/sites", headers = headers).status_code == 200:
            print("    " + "\"sites_read_all\": " + "true" + ",")
        else:
            print("    " + "\"sites_read_all\": " + "false" + ",")
        # Permission: User.Read.All
        if requests.get(r"https://graph.microsoft.com/v1.0/users", headers = headers).status_code == 200:
            print("    " + "\"user_read_all\": " + "true" + "")
        else:
            print("    " + "\"user_read_all\": " + "false" + "")
        print("}")
    except:
        pass

## Process
# Loop
for _ in range(3):
    # Call main()
    main()
