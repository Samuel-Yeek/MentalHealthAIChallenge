import pandas as pd
from datetime import datetime
import win32com.client as win32

# Read Excel sheets
users_df = pd.read_excel('data.xlsx', sheet_name='users')
messages_df = pd.read_excel('data.xlsx', sheet_name='messages')

# Get current date and day of year
current_date = datetime.now()
formatted_date = current_date.strftime("%d-%b")  # Format date as DD-MonthName for display
day_of_year = current_date.timetuple().tm_yday  # Get day of year (1-365)

# Find the message, riddle, and answer for today's day of year
today_message_row = messages_df[messages_df['date'] == day_of_year]

if today_message_row.empty:
    print(f"No message found for day {day_of_year}. Please check your messages sheet.")
    exit()

message = today_message_row['message'].values[0]
riddle = today_message_row['riddle'].values[0]
yesterday_riddle_answer = today_message_row['yesterdayriddleanswer'].values[0]

# Create Outlook application object
outlook = win32.Dispatch('outlook.application')

# Send emails to users
for _, user in users_df.iterrows():
    recipient_email = user['email']
    firstname = user['FirstName']
    lastname = user['LastName']
    
    # Create email
    mail = outlook.CreateItem(0)
    mail.To = recipient_email
    mail.Subject = f'Your Daily Encouragement - {formatted_date}'
    mail.Body = (
        f"Dear {firstname} {lastname},\n\n"
        f"Today's message ({formatted_date}):\n\n{message}\n\n"
        f"Riddle of the Day:\n{riddle}\n\n"
        f"Yesterday's Riddle Answer:\n{yesterday_riddle_answer}\n\n"
        f"Best regards,\nYour Friend Sam's Bot"
    )
    
    # Send email
    mail.Send()

print(f"Encouraging emails for {formatted_date} (Day {day_of_year}) sent successfully!")