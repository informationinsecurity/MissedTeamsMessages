import requests
import json
import re
from datetime import timedelta
from datetime import datetime, timezone
import arrow
import pandas as pd
import time
from slack_webhook import Slack
import unicodedata
from slack_sdk import WebClient

#runs dev features for troubleshooting
DEV = 0

enumerate_teams = 1
#timeout in mins
TIMEOUT = 10

#Alert to slack channel when missed message is found
SEND_TO_SLACK = 0


azure_tenant_id = "<Azure tenant id here>"
azure_client_id = "<Azure_client_id here>"
azure_client_secret = "<Azure_client_secret here>"

TIMEOUT = TIMEOUT * 60

TIMEOUT_MINS = TIMEOUT / 60

historic_replies = ''
historic_messages = ''
api_call_count = 0
total_api_call_count = 0
run_count = 0
team_search_count = 0

if DEV == 0:
	#webhook to post to
    slack_webhook = '<SLACK_WEBHOOK>'
    #channel to post to - for replies using api instead of webhook
	slack_channel_id = "<SLACK_CHANNEL_ID>"
    #Team Members to exclude from alerting (your team messaged not customer or contact)
	team_name = '<firstname lastname>', <firstname lastname>'

if DEV == 1:
	#For dev I exclude myself from team to trigger alerts for testing
    #webhook to post to
    slack_webhook = '<SLACK_WEBHOOK>'
    #channel to post to - for replies using api instead of webhook
	slack_channel_id = "<SLACK_CHANNEL_ID>"
    #Team Members to exclude from alerting (your team messaged not customer or contact)
	team_name = '<firstname lastname>', <firstname lastname>'

###-Slack Token Info#####
slack_bot_token = '<Slack_Bot_Token XOXB->'
slack_signing_secret = '<slack_signing_secret>'
client = WebClient(token=slack_bot_token)

# Where to post / pull history from in slack - ex: #missed-messages
#slack_channel_id = "<Channel ID>"


total_api_call_count = 0

while True:
    try:
        team_search_count = 0
        run_count += 1
        teams_list = ''
        api_call_count = 0
        #### Graph API Auth ######
        # Get  Token for Graph API
        url = "https://login.microsoftonline.com/"+str(azure_tenant_id)+"/oauth2/v2.0/token"
        payload = str(azure_client_id) + "&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default&client_secret="+str(azure_client_secret)+"&grant_type=client_credentials"
		
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }
        response = requests.request("POST", url, headers=headers, data=payload)
        json_resp = response.json()
        access_token = 'Bearer ' + json_resp['access_token']

        headers = {
            "Authorization": access_token
        }

    except:
        print("Unable to Pull Auth Token")
        pass

    try:
        ### Pull Teams List
        result = requests.get(f'https://graph.microsoft.com/beta/teams', headers=headers)
        result.raise_for_status()
        teams_json_result = result.json()
        dict_object = json.dumps(teams_json_result, indent = 4)
        next_page_link = teams_json_result['@odata.nextLink']
        #print(next_page_link)
        #print(dict_object['value']['id'])
        team_count = 0
        api_call_count += 1
        total_api_call_count += 1
    except:
        print("Unable to Pull Teams List")
        pass

    for item in dict_object.split("\"id\":"):
        try:
            #print(item)
            team_id = re.search(r"(.*?,)",item)
            team_id = team_id.group(1)
            team_id = str(team_id).replace(' \"','')
            team_id = str(team_id).replace('\",','')
            #print(team_id)
            team_name = re.search(r"(displayName.*?,)",item)
            team_name = team_name.group(1)
            team_name = str(team_name).replace('displayName\": \"','')
            team_name = str(team_name).replace('\",','')
            #print(team_name)
			team_count += 1
			teams_list += team_id + "\n"

        except:
            print("Unable to parse Teams List")
            pass
    
    try:
        # print(teams_list)
        print("Total Teams Identified: " + str(team_count))
        print("--------------------")

    except:
        print("Unable to print Teams List")
        pass
    ### Pull Teams Channel List ###
    for team in teams_list.split():
        td_in_hours = 0
        td_in_minutes = 0
        reply_td_in_minutes = 0
        reply_td_in_hours = 0
        message_detail = ''

        #print("Checeking Channel list for Team ID: " + str(team))

        try:
            channel_url = 'https://graph.microsoft.com/v1.0/teams/'+team+'/channels'
            response = requests.request("GET", channel_url, headers=headers, data=payload)
            channel_json_result = response.json()
            api_call_count += 1
            total_api_call_count += 1
        except:
            print("Unable to pull Channel List")
            pass

        try:
            ### Parse Channel IDs
            channel_id = re.search(r"(\[{\'id\':.*?,)", str(channel_json_result))
            channel_id = channel_id.group(1)
            channel_id = str(channel_id).replace('[{\'id\': \'', '')
            channel_id = str(channel_id).replace('\',', '')
        except:
            print("Unable to parse Channel List")
            pass

        try:
            ### Pull Team Name
            team_lookup = 'https://graph.microsoft.com/v1.0/teams/' + team + '/channels/' + channel_id
            response = requests.request("GET", team_lookup, headers=headers, data=payload)
            team_lookup_result = response.json()
            team_lookup_name = team_lookup_result['description']
            team_lookup_channel_name = team_lookup_result['displayName']
            #print("Checking Messages - Team: " + str(team_lookup_name) + " and Channel: " + str(team_lookup_channel_name))
            api_call_count += 1
            total_api_call_count += 1
        except:
            print("Unable to Pull Team Name")
            pass

        try:
            ### Pull Message for Team
            message_url = 'https://graph.microsoft.com/beta/teams/' + str(team) + '/channels/' + str(channel_id) + '/messages'
            message_response = requests.request("GET", message_url, headers=headers, data=payload)
            message_json_result = message_response.json()
            json_object = json.dumps(message_json_result, indent=4)
            api_call_count += 1
            total_api_call_count += 1
            team_search_count += 1


        except:
            print("Unable to Pull Messages")
            pass

        try:
            if message_json_result['value']:
                #try:
                latest_message_object = message_json_result['value'][0]
                # print(latest_message_object)

                message_from = ''
                message_id = latest_message_object['id']
                message_type = latest_message_object['messageType']
                message_create_date = latest_message_object['createdDateTime']
                message_reactions = latest_message_object['reactions']
                message_content = latest_message_object['body']['content']
                try:
                    message_content_short = message_content[:30]
                except:
                    message_content_short = ''
                try:
                    message_deleted = latest_message_object['deletedDateTime']
                except:
                    message_deleted = ''
                if message_deleted:
                    print("Message was deleted!")
                
                try:
                    message_from = latest_message_object['from']['user']['displayName']
                except:
                    pass


                try:
                    ### Get Time Delta ####
                    # teams_timestamp = "2023-06-07T06:39:48.193Z"
                    # Current Time:     2023-06-07T151313.000Z
                    now = datetime.utcnow()
                    current_time = now.strftime("%Y-%m-%dT%H:%M:%S.000Z")
                    message_teams_date = arrow.get(message_create_date)
                    current_date = arrow.get(current_time)
                    time_diff = current_date - message_teams_date
                    time_delta = pd.Timedelta(time_diff)
                    total_seconds = time_delta.total_seconds()
                    seconds_in_hour = 60 * 60
                    td_in_hours = total_seconds / seconds_in_hour
                    td_in_minutes = total_seconds / 60

                    # number of hours since last reply
                    message_hours_since_message = td_in_hours
                    message_mins_since_message = td_in_minutes
                except:
                    print("Unable to calculate Time difference")
                    pass

                if not message_from:
                    message_from = "Bot"
                try:
                    message_detail = "Count: " + str(team_search_count) + "\nTeam: " + str(team_lookup_name) + "\nChannel: " + str( team_lookup_channel_name) + "\nElapsed Time: " + str(round(message_mins_since_message,2)) + " minutes \nFrom: " + str(message_from) + "\nLatest Message: " + str(message_content_short) + "\nMessage ID: " + str(message_id)
                except:
                    print("Unable to set message detail")
                    pass

                if message_reactions:
                    message_detail += "\nMessage Reaction: Yes"

            elif not message_json_result['value']:
                pass
        except KeyError:
            print("Error getting value")
            pass

        try:
            reply_url = 'https://graph.microsoft.com/beta/teams/' + str(team) + '/channels/' + str(channel_id) + '/messages/' + str(message_id) + '/replies'
            message_response = requests.request("GET", reply_url, headers=headers, data=payload)
            reply_json_result = message_response.json()
            reply_json_object = json.dumps(reply_json_result, indent=4)
            api_call_count += 1
            total_api_call_count += 1

        except:
            print("Unable to pull replies")
            pass
        try:
            if reply_json_result['value']:
                try:
                    latest_reply_object = reply_json_result ['value'][0]
                    # print(reply_json_result)
                    reply_message_from = ''
                    reply_message_id = latest_reply_object['id']
                    reply_message_type = latest_reply_object['messageType']
                    reply_message_create_date = latest_reply_object['createdDateTime']
                    reply_message_reactions = latest_reply_object['reactions']
                except:
                    print("Unable to parse replies")
                    pass

                try:
                    reply_message_from = latest_reply_object['from']['user']['displayName']
                except:
                    pass

                try:
                    repy_message_deleted = latest_reply_object['deletedDateTime']

                except:
                    repy_message_deleted = ''

                reply_message_content = latest_reply_object['body']['content']
                reply_message_content_short = reply_message_content[:30]

                try:
                    now = datetime.utcnow()
                    current_time = now.strftime("%Y-%m-%dT%H:%M:%S.000Z")

                    reply_message_teams_date = arrow.get(reply_message_create_date)
                    current_date = arrow.get(current_time)

                    # time_diff = current_date - message_teams_date
                    reply_time_diff = current_date - reply_message_teams_date

                    reply_time_delta = pd.Timedelta(reply_time_diff)

                    reply_total_seconds = reply_time_delta.total_seconds()
                    seconds_in_hour = 60 * 60
                    reply_td_in_hours = reply_total_seconds / seconds_in_hour
                    reply_td_in_minutes = reply_total_seconds / 60

                    # number of hours since last reply
                    reply_message_hours_since_message = reply_td_in_hours
                    reply_message_mins_since_message = reply_td_in_minutes

                except:
                    print("Unable to calculate reply times")
                    pass

                message_detail += "\nReply From: " + str(reply_message_from) + "\nElapsed Reply Time: " +str(round(reply_message_mins_since_message,2)) + "\nLatest Reply: " + str(reply_message_content_short)

                if reply_message_reactions:
                    message_detail += "\nReply Reaction: True"

            elif not reply_json_result['value']:
                reply_message_content = ''

        except KeyError:
            print("Error getting value")
            pass




        try:
            fire_webhook = 0
            if message_from not in team_name and message_from != "Bot" and not message_deleted and not reply_message_content and not message_reactions and message_mins_since_message >= 10 and message_hours_since_message <= 12:
                fire_webhook = 1
                print("Firing Webhook")

            if reply_message_content and reply_message_from not in team_name and not repy_message_deleted and not reply_message_reactions and reply_message_mins_since_message > 10 and reply_message_hours_since_message <= 12:
                fire_webhook = 1
                print("Firing Webhook")

        except:
            print("Unable to trigger webhook!")
            pass

        # try:
        if fire_webhook == 1:
            # conversations.history returns the first 100 messages by default
            result = client.conversations_history(channel=slack_channel_id)
            conversation_history = result["messages"]
            message_found = 0
            slack_reaction = False
            matching_message_ts = ''

            for individual_message in conversation_history:
                single_message = individual_message["text"]
                single_message_ts = individual_message["ts"]
                reply_counter = 0
                try:
                    reply_counter = individual_message['reply_count']
                except:
                    pass

                if message_id in single_message:
                    if "reactions" in individual_message:
                        slack_reaction = True

                    if slack_reaction == False :
                        print("Found Prior Message regarding this message without reaction- tagging DRAMS")
                        print("reply count: " + str(reply_counter))
                        slack_message = message_detail
                        result = client.chat_postMessage(channel=slack_channel_id, thread_ts=single_message_ts, text=slack_message)
                        message_found = 1
                        matching_message_ts = single_message_ts
                        if int(reply_counter) >= 2:
                           
                            print("BLOWING UP THE TEAM")
                            slack_message = "<!subteam^groupid_here> - No one has responded to the message in Teams Channel!:" + str(team_lookup_channel_name) + "> in over " + str(round(message_mins_since_message, 2)) + " minutes\n" + message_detail + "\nClear this message by replying or reacting  in <#channel_id_here>"
                            casenotify_slack = Slack(url=team_webhook_here)
                            casenotify_slack.post(text=slack_message)

            if message_found == 0 and slack_reaction == False:
                print("Message was not found in history - sending new message")
                slack_message = message_detail
                casenotify_slack = Slack(url=slack_webhook)
                casenotify_slack.post(text=slack_message)




            message_found = 0

        
        try:
            print(message_detail)
            print("--------------------")
        except:
            print("Unable to print details")
            pass
    print(str(team_search_count) + " Teams Channels searched")
    print(str(run_count) + " Total Runs")
    print(str(api_call_count) + " API Calls this round")
    print(str(total_api_call_count) + " API Calls in total")
    print("Scan Finished! Sleeping for: " +str(TIMEOUT_MINS) + " minutes")
    time.sleep(TIMEOUT)
