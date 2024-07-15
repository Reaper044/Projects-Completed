# This is a code to send excel file in .xlsx format to outlook mail as attachment

#Libraries
from metabasepy import Client
import pandas as pd
from io import BytesIO
import requests
import re
import json

script_name = 'your_script_name'

# Add the mail ID to which the data is to be sent
# You can also replace None in cc and bcc with any email ID or multiple ID's within [""]
# Ex: ["abc@gmail.com","xyz@gmail.com"]
mail_recipient_list = ["xyz@gmail.com"]
mail_cc_list = None
mail_bcc_list = None

# Enter the attachments in the same format as done below
# You can add a new attachment by typing its card ID for the question this can be found from the url
# For ex from
# https://metabase.protium.co.in/question/123456-daily-mis-raw-data
# Take the number mentioned just before the file name, so the card id to be entered will be 10485
# If 'None' is provided no column operations are done on data
# Ex: 10101: None
# After the card ID put :
# Within in 'columns' [] enter the column names from query in the same sequence in which you want the column names within the attachment
# In date_columns write all column names which contains dates as values
questions_to_columns = {
    11111: {
        'columns': [
            'Lead Code', 'Case Name', 'Branch', 'Cluster', 'Product', 'SO ID',
            'Lead Created Date', 'login_date', 'Status', 'Current Stage', 'Case Rejection Date',
            'CM of Case', 'Send For Approval Amount', 'Approved Amount', 'Disbursed Amount',
            'Loan Number', 'Eligible Program', 'Send For Approval Date', 'Approval Date',
            'Disbursed Date', 'Fundout ROI (%)', 'Insurance Premium Amount During Fundout',
            'fundout PF (%)', 'Push to LMS Date', 'Push to LMS Amount', 'Push to LMS ROI',
            'Push to LMS PF', 'insurance_premium_amount during push to lms stage'
        ],
        'date_columns': [
            'Lead Created Date', 'login_date', 'Case Rejection Date',
            'Send For Approval Date', 'Approval Date',
            'Disbursed Date', 'Push to LMS Date'],
        'number_column' : ['Send For Approval Amount','fundout PF (%)','Push to LMS Amount','Push to LMS ROI','Fundout ROI (%)',
                'Push to LMS PF',
                'Insurance Premium Amount During Fundout',
                'Approved Amount',
                'Disbursed Amount','insurance_premium_amount during push to lms stage']
    },
    22222: {
        'columns': [
            'Lead Code', 'Case Name', 'Branch', 'Cluster', 'Product', 'SO ID',
            'Lead Created Date', 'login_date', 'Status', 'Current Stage', 'Case Rejection Date',
            'CM of Case', 'Send For Approval Amount', 'Approved Amount', 'Disbursed Amount',
            'Loan Number', 'Eligible Program', 'Send For Approval Date', 'Approval Date',
            'Disbursed Date', 'Fundout ROI (%)', 'Insurance Premium Amount During Fundout',
            'fundout PF (%)', 'Push to LMS Date', 'Push to LMS Amount', 'Push to LMS ROI',
            'Push to LMS PF', 'insurance_premium_amount during push to lms stage'
        ],
        'date_columns': [
            'Lead Created Date', 'login_date', 'Case Rejection Date',
            'Send For Approval Date', 'Approval Date',
            'Disbursed Date', 'Push to LMS Date'],
        'number_column': ['Send For Approval Amount', 'Approved Amount', 'Disbursed Amount',
                          'Fundout ROI (%)', 'Insurance Premium Amount During Fundout', 'fundout PF (%)',
                          'Push to LMS Amount', 'Push to LMS ROI', 'Push to LMS PF', 'insurance_premium_amount during push to lms stage']

    },
    33333: {
        'columns': [
                    'Lead Code', 'Case Name', 'Branch', 'Cluster', 'Product', 'SO ID', 'DSA Code',
                    'Lead Created Date', 'Login Date', 'Status', 'Current Stage', 'Case Rejection Date',
                    'Send For Approval Amount', 'Approved Amount', 'Disbursed Amount', 'Loan Number',
                    'Eligible Program', 'Send For Approval Date', 'Approval Date', 'Disbursed Date',
                    'Fundout ROI (%)', 'Insurance Premium Amount During Fundout', 'fundout PF (%)',
                    'Push to LMS Date', 'Push to LMS Amount', 'Push to LMS ROI', 'Push to LMS PF',
                    'insurance_premium_amount during push to lms stage','Applied Loan Amount'
        ],
        'date_columns': [
                    'Lead Created Date', 'Login Date', 'Case Rejection Date',
                    'Send For Approval Date', 'Approval Date', 'Disbursed Date',
                    'Push to LMS Date'],
        'number_column' :[
                    'Send For Approval Amount', 'Approved Amount', 'Disbursed Amount',
                    'Fundout ROI (%)', 'Insurance Premium Amount During Fundout', 'fundout PF (%)',
                    'Push to LMS Amount', 'Push to LMS ROI', 'Push to LMS PF','insurance_premium_amount during push to lms stage','Applied Loan Amount'
        ]
    },
    44444: {
        'columns': [
                    'Lead Code', 'Case Name', 'Branch', 'Cluster', 'Product', 'SO ID', 'DSA Code',
                    'Lead Created Date', 'Login Date', 'Status', 'Current Stage', 'Case Rejection Date',
                    'Send For Approval Amount', 'Approved Amount', 'Disbursed Amount', 'Loan Number',
                    'Eligible Program', 'Send For Approval Date', 'Approval Date', 'Disbursed Date',
                    'Fundout ROI (%)', 'Insurance Premium Amount During Fundout', 'fundout PF (%)',
                    'Push to LMS Date', 'Push to LMS Amount', 'Push to LMS ROI', 'Push to LMS PF',
                    'insurance_premium_amount during push to lms stage','Applied Loan Amount'
        ],
        'date_columns': [
                        'Lead Created Date', 'Login Date', 'Case Rejection Date',
                        'Send For Approval Date', 'Approval Date', 'Disbursed Date',
                        'Push to LMS Date'],
        'number_column': [                    
                    'Send For Approval Amount', 'Approved Amount', 'Disbursed Amount',
                    'Fundout ROI (%)', 'Insurance Premium Amount During Fundout', 'fundout PF (%)',
                    'Push to LMS Amount', 'Push to LMS ROI', 'Push to LMS PF',
                    'insurance_premium_amount during push to lms stage','Applied Loan Amount']
    },
    55555: {
        'columns': [
                    'Lead Code', 'Case Name', 'Branch', 'Cluster', 'Product', 'SO ID', 'DSA Code',
                    'Lead Created Date', 'login_date', 'Status', 'Current Stage', 'Case Rejection Date',
                    'Send For Approval Amount', 'Approved Amount', 'Disbursed Amount', 'Loan Number',
                    'Eligible Program', 'Send For Approval Date', 'Approval Date', 'Disbursed Date',
                    'Fundout ROI (%)', 'Insurance Premium Amount During Fundout', 'fundout PF (%)',
                    'Push to LMS Date', 'Push to LMS Amount',
                     'Push to LMS ROI', 'Push to LMS PF',
                    'insurance_premium_amount during push to lms stage'
        ],
        'date_columns': [
                        'Lead Created Date', 'login_date', 'Case Rejection Date',
                        'Send For Approval Date', 'Approval Date', 'Disbursed Date',
                        'Push to LMS Date'],
        'number_column':[
                    'Send For Approval Amount', 'Approved Amount', 'Disbursed Amount',
                    'Fundout ROI (%)', 'Insurance Premium Amount During Fundout', 'fundout PF (%)',
                     'Push to LMS Amount', 'Push to LMS ROI', 'Push to LMS PF',
                    'insurance_premium_amount during push to lms stage']
    },
    66666: {
        'columns': [
                    'Lead Code', 'Case Name', 'Branch', 'Cluster', 'Product', 'SO ID', 'Lead Created Date',
                    'login_date', 'Status', 'Current Stage', 'Case Rejection Date', 'CM of Case',
                    'Send For Approval Amount', 'Approved Amount', 'Disbursed Amount', 'Loan Number',
                    'Eligible Program', 'Send For Approval Date', 'Approval Date', 'Disbursed Date',
                    'Fundout ROI (%)', 'Insurance Premium Amount During Fundout', 'fundout PF (%)',
                    'Push to LMS Date', 'Push to LMS Amount', 'Push to LMS ROI', 'Push to LMS PF',
                    'insurance_premium_amount during push to lms stage'
        ],
        'date_columns': [
                    'Lead Created Date', 'login_date',
                    'Case Rejection Date', 'Send For Approval Date', 'Approval Date', 'Disbursed Date', 'Push to LMS Date'
                    ],
        'number_column':['Send For Approval Amount', 'Approved Amount', 'Disbursed Amount',
                    'Fundout ROI (%)', 'Insurance Premium Amount During Fundout', 'fundout PF (%)',
                    'Push to LMS Amount', 'Push to LMS ROI', 'Push to LMS PF',
                    'insurance_premium_amount during push to lms stage']
    }
}

# Write the name which you want for your attachment within the mail
question_to_script_name = {
                            11111: 'EEG_Raw_Data', 22222: 'BBL_Raw_Data',
                            33333: 'SME_Raw_Data', 44444: 'SME_BL_Raw_Data',
                            55555: 'MF_Raw_Data', 66666: 'EIL_Raw_Data'
                        }

team_name = 'SBL'


def fetch_data(card_id):
# Enter Your Email ID and Password in the quotes "abc@protium.co.in", password="abc@12"
    cli = Client(username="prerak.joshi@gmail.com", password="abc@123",
                 base_url="https://metabase.protium.co.in")
    cli.authenticate()
    json_result = cli.cards.download(card_id=card_id, format='json')
    if isinstance(json_result, dict):
        error = json_result.get('error')
        raise Exception(f'Error in Metabase question {card_id}: {error}')
    df_bre_alerts = pd.DataFrame(json_result)
    
    if card_id in questions_to_columns:
        columns = questions_to_columns[card_id]['columns']
        date_columns = questions_to_columns[card_id]['date_columns']
        number_columns = questions_to_columns[card_id]['number_column']
        for col in date_columns:
            df_bre_alerts[col] = pd.to_datetime(df_bre_alerts[col], errors='coerce', format = '%d-%m-%Y, %I:%M %p').dt.tz_localize(None)
        for col2 in number_columns:
            df_bre_alerts[col2] = df_bre_alerts[col2].apply(lambda x:float(x.replace(",","")) if pd.notnull(x) else None)
    return df_bre_alerts


# If anyone is present in cc or bcc write their email ID's in cc_list and bcc_list
# Change contents or subject of mail from below
def auto_mail(file_name='attachment',
              df_dict=None,
              recipient_list=None,  # list
              cc_list=None,  # list
              bcc_list=None,  # list
              subject="Subject",
              body_greeting="Hi,",
              body_content="This is a test email",
              body_closing="Regards,",
              body_sender_name="Team Protium",
              los_app_id=script_name,
              los_app_entity_id=team_name,
              los_type='con_fin'
              ):
    file_op = BytesIO()
    if df_dict is not None:
        if not isinstance(df_dict, dict):
            df_dict = {'data': df_dict}
        writer = pd.ExcelWriter(file_op, engine='xlsxwriter')
        for df in df_dict:
            df_dict[df].to_excel(writer, sheet_name=df)
        writer.close()
        file_name = file_name
    file_data = file_op.getvalue()

    url = "https://api-marketplace.prod.growth-source.com:443/api/24057813-communication/operation" \
          "/sendEmailWithAttachments/execute/communication/send-email-with-attachments"

    email_body_html = (
        f""" {body_greeting}

        {body_content}

        {body_closing}
        {body_sender_name}

        """.replace('\n', '<br>'))

    request_dict = {
        "from": "no-reply@protium.co.in",
        "losAppId": los_app_id,
        "losAppEntityId": los_app_entity_id,
        "losType": los_type,
        "subject": subject,
        "bodyHtml": email_body_html,
        "recipientList": recipient_list,
        "ccList": cc_list,
        "bccList": bcc_list
    }
    payload = {'requestJson': json.dumps(request_dict)}

    files = [
        ('attachments', (file_name, file_data, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))]
    headers = {
        'accept': 'application/json'
    }
    response = requests.request("POST", url, headers=headers, data=payload, files=files)
    return response, response.text


def take_data_from_metabase_and_send_email(metabase_question):
    try:
        data = fetch_data(metabase_question)

        if questions_to_columns[metabase_question] is not None:
            data = data[questions_to_columns[metabase_question]['columns']]
# The attachment name is written in the line just below
# Add anything to its name by writing + and add a suffix within ""
# For a prefix (Writing before the attachment name) write it just after the = within "" and write + after closing the "
        auto_mail(file_name=question_to_script_name[metabase_question] + ".xlsx",
                  df_dict={'data': data},
                  recipient_list=mail_recipient_list,
                  cc_list=mail_cc_list,
                  bcc_list=mail_bcc_list,
                  subject=question_to_script_name[metabase_question],
# Below is the body of the mail where name of the file is dynamically attached
                  body_greeting="Dear all,",
                  body_content=f"Please find the attached {question_to_script_name[metabase_question]}.")

    except Exception as ex:
        auto_mail(df_dict=None,
                  recipient_list=mail_cc_list,
                  subject=f"Failure in {script_name} Code",
                  body_greeting="Hi team, ",
                  body_content=f"Error occurred in {script_name}: {ex}")
        raise Exception(f"Error Occurred in {script_name}: {ex}")


if __name__ == "__main__":
    for question in questions_to_columns:
        take_data_from_metabase_and_send_email(metabase_question=question)
