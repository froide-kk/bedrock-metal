import requests
import openpyxl
from dotenv import load_dotenv
import os
from datetime import datetime
import pytz

def fetch_messages(token, channel_id):
    url = "https://slack.com/api/conversations.history"
    headers = {"Authorization": f"Bearer {token}"}
    params = {"channel": channel_id}
    all_messages = []

    while True:
        response = requests.get(url, headers=headers, params=params).json()
        if not response['ok']:
            raise Exception("Error fetching messages:", response.get('error'))
        
        messages = response['messages']
        all_messages.extend(messages)

        if not response.get('has_more'):
            break

        params['cursor'] = response['response_metadata']['next_cursor']

    return all_messages

def convert_timestamp(ts):
    # UTC時刻に変換し、JSTに変換する
    utc_time = datetime.utcfromtimestamp(float(ts))
    jst_time = utc_time.replace(tzinfo=pytz.utc).astimezone(pytz.timezone("Asia/Tokyo"))
    return jst_time.strftime("%Y/%m/%d %H:%M:%S")

def save_messages_to_excel(messages, file_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["User ID", "Timestamp (JST)", "Text"])  # Excelのヘッダー

    for msg in messages:
        timestamp_jst = convert_timestamp(msg.get('ts'))
        ws.append([msg.get('user'), timestamp_jst, msg.get('text')])

    wb.save(file_path)

def main():
    load_dotenv()
    slack_token = os.getenv('SLACK_TOKEN')
    channel_id = os.getenv('CHANNEL_ID')

    if not slack_token or not channel_id:
        print("Error: Please set SLACK_TOKEN and CHANNEL_ID in your .env file.")
        return

    print("Fetching messages from Slack...")
    messages = fetch_messages(slack_token, channel_id)
    print(f"Total {len(messages)} messages fetched.")

    file_path = 'slack_messages.xlsx'
    save_messages_to_excel(messages, file_path)
    print(f"Messages saved to {file_path}.")

if __name__ == "__main__":
    main()

