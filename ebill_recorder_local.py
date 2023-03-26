from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.auth.exceptions import RefreshError
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build
from datetime import datetime, timedelta
from base64 import urlsafe_b64decode
import os.path
import gspread
import quopri
import email
import time
import sys
import re


# If modifying these scopes, delete the file token.json.
SHEET_NAME = '驊驊愛記帳 3.0'
WORKSHEET_NAME = '外送'
TOKEN_PATH = 'token.json'
CREDENTIALS_PATH = 'credentials.json'
SINCE_DATE = "2022/06/15"  # 最早可追溯日期
SCOPES = ['https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/spreadsheets', 
          'https://www.googleapis.com/auth/gmail.readonly']


class EbillRecorder:
    def __init__(self, subject, date_split_symbol, date_col,
                 pattern_restaurant, pattern_date, pattern_cost):
        self.subject = subject
        self.date_split_symbol = date_split_symbol
        self.date_col = date_col
        self.pattern_restaurant = pattern_restaurant
        self.pattern_date = pattern_date
        self.pattern_cost = pattern_cost
        self.creds = self.get_credentials()
        start_row, start_date = self.find_start_point()
        self.record_latest_info_from_ebil(start_row, start_date)

    def get_credentials(self):
        """設置認證，從文件中讀取或要求用戶進行認證，並將結果保存。"""
        try:
            creds = None
            # The file token.json stores the user's access and refresh tokens, and is
            # created automatically when the authorization flow completes for the first
            # time.
            if os.path.exists(TOKEN_PATH):
                creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
            # If there are no (valid) credentials available, let the user log in.
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        CREDENTIALS_PATH, SCOPES)
                    creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open(TOKEN_PATH, 'w') as token_file:
                    token_file.write(creds.to_json())
        except RefreshError:
            os.remove('token.json')
            return self.get_credentials()
        return creds
    
    def find_start_point(self):
        # retrive obj from Google Sheet
        client = gspread.authorize(self.creds)
        self.sheet = client.open(SHEET_NAME).worksheet(WORKSHEET_NAME)
        # find last point
        values = self.sheet.col_values(self.date_col)
        last_row = len(values)
        last_date_string = values[-1] if values and values[-1]!='日期' else SINCE_DATE
        # compute start point
        start_row = last_row + 1
        last_date_obj = datetime.strptime(last_date_string, r"%Y-%m-%d")
        start_date_obj = last_date_obj + timedelta(days=1)
        start_date_string = start_date_obj.strftime("%Y-%m-%d")
        return start_row, start_date_string
    
    def record_latest_info_from_ebil(self, start_row, start_date):
        """
        透過Gmail API建立連線，取得符合條件的信件，
        再解析以取得相關資訊。最後將解析出的資訊印出。
        """
        self.row = start_row
        try:
            # Call the Gmail API
            service = build('gmail', 'v1', credentials=self.creds)
            # search
            query = f"subject:{self.subject} after:{start_date}"
            results = service.users().messages().list(userId='me', q=query).execute()
            messages = results.get('messages', [])
            if not messages:
                print('No messages found.')
                return
            # Main loop
            for message in reversed(messages):
                msg = service.users().messages().get(
                    userId='me', id=message['id'], format='raw').execute()
                self.processing(msg['raw'])
                self.row += 1
        except HttpError as error:
            # TODO(developer) - Handle errors from gmail API.
            print(f'An error occurred: {error}')

    def processing(self, raw_msg):
        """解析郵件並將其解碼為純文本，然後使用正則表達式匹配並提取餐廳名稱、日期和金額。"""
        mime_msg = self.retrive_desired_mime_msg(raw_msg)
        decoded_msg = self.decode_qp(mime_msg)
        cleaned_msg = self.remove_tags(decoded_msg)
        info = self.extract_info(cleaned_msg)
        self.write_to_sheet(info)
        print(info)

    def retrive_desired_mime_msg(self, msg):
        """需要在子類依不同信件結構實現"""
        return None

    def decode_qp(self, string_to_convert):
        """將Quoted-Printable格式的字串解碼，去除標籤並以正則表達式擷取印出。"""
        bytes_to_convert = quopri.decodestring(string_to_convert.encode('utf-8'))
        decoded_string = bytes_to_convert.decode('utf-8')
        return decoded_string

    def remove_tags(self, decoded_string):
        """將html中的標籤去除以便抽取資訊"""
        cleaned_string = re.sub('<[^>]+>', '', decoded_string)
        return cleaned_string

    def extract_info(self, cleaned_string):
        """以正則表達式從解碼後的字串中擷取餐廳名稱、日期和金額，並以特定格式輸出。"""
        # retaurant       
        match_restaurant = re.search(self.pattern_restaurant, cleaned_string)
        restaurant = match_restaurant.group(1) if match_restaurant else ''
        # date
        match_date = re.search(self.pattern_date, cleaned_string)
        date = match_date.group(0) if match_date else ''
        year, month, day = date.split(self.date_split_symbol)
        formatted_date = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
        # cost
        match_cost = re.search(self.pattern_cost, cleaned_string)
        cost = match_cost.group(1) if match_cost else ''
        return [restaurant, formatted_date, cost]

    def write_to_sheet(self, info):
        # TODO record one row once a time
        # A = 65
        # left_chr, right_chr = chr(64 + self.date_col - 1), chr(64 + self.date_col + 1)
        # range_name = f"{left_chr}{self.row}:{right_chr}{self.row}"
        # formated_info = [i for i in info]
        # self.sheet.update(range_name, formated_info)

        # positions = [chr(64 + self.date_col + i)+str(self.row) for i in range(-1, 2)]
        # info_dict = dict(zip(positions, info))
        # self.sheet.update(info_dict)

        # record one cell a time
        rest, date, cost = info
        try:
            self.sheet.update_cell(self.row, self.date_col-1, rest)
            self.sheet.update_cell(self.row, self.date_col, date)
            self.sheet.update_cell(self.row, self.date_col+1, cost)  
        except gspread.exceptions.APIError:
            self.count_down()
            self.write_to_sheet(info)

    def count_down(self):
        print('Waiting for API coldtime')
        for i in range(60, 0, -1):
            sys.stdout.write(f'\r{i}')
            sys.stdout.flush()
            time.sleep(1)
        print()

class UbereatsEbillRecorder(EbillRecorder):
    def __init__(self):
        super().__init__(
            subject= r"透過 Uber Eats 系統送出的訂單",
            pattern_restaurant = r"以下是您在(.+)訂購的電子明細。",
            pattern_date = r"\d{4}/\d{1,2}/\d{1,2}",
            pattern_cost = r"總計\s*\$\s*([\d,]+)",
            date_split_symbol = r"/",
            date_col = 5
        )
    
    def retrive_desired_mime_msg(self, msg):
        """取得Ubereats的MIME格式資訊"""
        mime_msg = email.message_from_bytes(urlsafe_b64decode(msg))
        return mime_msg.get_payload()


class FoodpandaEbillRecoder(EbillRecorder):
    def __init__(self):
        super().__init__(
            subject= r"你的訂單已成功訂購",
            pattern_restaurant = r"我們已收到[你您]在 (.+) 下訂?的訂單囉！", 
            pattern_date = r"\d{4}-\d{1,2}-\d{1,2}",
            pattern_cost = r"訂單總額\s*\$\s*([\d,]+)",
            date_split_symbol = r"-",
            date_col = 2
        )
    
    def retrive_desired_mime_msg(self, msg):
        """取得Foodpanda的MIME格式資訊"""
        mime_msg = email.message_from_bytes(urlsafe_b64decode(msg))
        desired_part = mime_msg.get_payload()[1]  # 取第2個part
        return desired_part.get_payload()
    

if __name__ == '__main__':
    UbereatsEbillRecorder()
    FoodpandaEbillRecoder()