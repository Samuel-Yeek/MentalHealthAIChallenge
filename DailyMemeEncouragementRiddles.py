import pandas as pd
from datetime import datetime
import win32com.client as win32
import tkinter as tk
from tkinter import messagebox
import os

# Function to send emails
def send_emails():
    # Read Excel sheets
    users_df = pd.read_excel('data.xlsx', sheet_name='users')
    messages_df = pd.read_excel('data.xlsx', sheet_name='messages')

    # Get current date and day of year
    current_date = datetime.now()
    formatted_date = current_date.strftime("%d-%b")
    day_of_year = current_date.timetuple().tm_yday

    # Find today's message
    today_message_row = messages_df[messages_df['date'] == day_of_year]

    if today_message_row.empty:
        messagebox.showerror("Error", f"No message found for day {day_of_year}.")
        return

    message = today_message_row['message'].values[0]
    riddle = today_message_row['riddle'].values[0]
    yesterday_riddle_answer = today_message_row['yesterdayriddleanswer'].values[0]

    # Create Outlook application object
    outlook = win32.Dispatch('outlook.application')

    # Path to the image based on day of year
    image_path = os.path.join(os.getcwd(), 'Photos', f'{day_of_year}.png')

    # Send emails to users
    for _, user in users_df.iterrows():
        recipient_email = user['email']
        firstname = user['FirstName']
        lastname = user['LastName']

        # Create email
        mail = outlook.CreateItem(0)
        mail.To = recipient_email
        mail.Subject = f'Your Daily Meme and Encouragement - {formatted_date}'

        # HTML body with CSS styling and image alignment to the right, using CID for embedded image
        mail.HTMLBody = f"""
        <html>
            <head>
                <style>
                    body {{
                        font-family: Arial, sans-serif;
                        background-color: #f4f4f4;
                        margin: 20px;
                        padding: 20px;
                        border-radius: 5px;
                        box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                    }}
                    .container {{
                        display: flex;
                        align-items: center;
                    }}
                    .text {{
                        flex: 1;
                    }}
                    img {{
                        max-width: 200px; /* Adjust width as necessary */
                        margin-left: 20px; /* Space between text and image */
                    }}
                    h1 {{
                        color: #333;
                    }}
                    p {{
                        color: #555;
                    }}
                    .footer {{
                        margin-top: 20px;
                        font-size: small;
                        color: #888;
                    }}
                </style>
            </head>
            <body>
                <h1>Dear {firstname} {lastname},</h1>
                <div class="container">
                    <div class="text">
                        <p>Today's message ({formatted_date}):</p>
                        <p><strong>{message}</strong></p>
                        <p><strong>Riddle of the Day:</strong><br>{riddle}</p>
                        <p><strong>Yesterday's Riddle Answer:</strong><br>{yesterday_riddle_answer}</p>
                    </div>
                    <img src="cid:image" alt="Daily Image" />
                </div>
                <p class="footer">Best regards,<br>Your Friend Sam's Bot</p>
            </body>
        </html>
        """

        # Attach the image if it exists and set its CID for embedding in HTML
        if os.path.exists(image_path):
            attachment = mail.Attachments.Add(image_path)
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", "image")

        # Send email
        mail.Send()

    messagebox.showinfo("Success", f"Encouraging emails for {formatted_date} sent successfully!")

#Function to add a user to the Excel file
def add_user():
    first_name = entry_first_name.get()
    last_name = entry_last_name.get()
    email = entry_email.get()

    if not first_name or not last_name or not email:
        messagebox.showerror("Error", "All fields must be filled.")
        return

    # Read existing users and append new user
    users_df = pd.read_excel('data.xlsx', sheet_name='users')
    new_user = pd.DataFrame([[first_name, last_name, email]], columns=['FirstName', 'LastName', 'email'])
    
    updated_users_df = pd.concat([users_df, new_user], ignore_index=True)
    
    # Save back to Excel without altering other sheets
    with pd.ExcelWriter('data.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        updated_users_df.to_excel(writer, sheet_name='users', index=False)

    messagebox.showinfo("Success", "User added successfully!")
    
    # Clear input fields after adding user
    entry_first_name.delete(0, tk.END)
    entry_last_name.delete(0, tk.END)
    entry_email.delete(0, tk.END)

# Function to add a message for a specific day
def add_message():
    date_input = entry_date.get()
    message_input = entry_message.get()
    riddle_input = entry_riddle.get()
    yesterday_answer_input = entry_yesterday_answer.get()

    if not date_input or not message_input or not riddle_input or not yesterday_answer_input:
        messagebox.showerror("Error", "All fields must be filled.")
        return

    try:
        date_day_of_year = int(date_input)
        if date_day_of_year < 1 or date_day_of_year > 365:
            raise ValueError("Day must be between 1 and 365.")
        
        # Read existing messages and append new message
        messages_df = pd.read_excel('data.xlsx', sheet_name='messages')
        
        new_message_row = pd.DataFrame([[date_day_of_year, message_input, riddle_input, yesterday_answer_input]], 
                                        columns=['date', 'message', 'riddle', 'yesterdayriddleanswer'])
        
        updated_messages_df = pd.concat([messages_df, new_message_row], ignore_index=True)

        # Save back to Excel without altering other sheets
        with pd.ExcelWriter('data.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            updated_messages_df.to_excel(writer, sheet_name='messages', index=False)

        messagebox.showinfo("Success", "Message added successfully!")
        
        # Clear input fields after adding message
        entry_date.delete(0, tk.END)
        entry_message.delete(0, tk.END)
        entry_riddle.delete(0, tk.END)
        entry_yesterday_answer.delete(0, tk.END)

    except ValueError as e:
        messagebox.showerror("Error", str(e))

# Create the main window
root = tk.Tk()
root.title("User and Message Management")

# Create input fields and labels for adding a user
tk.Label(root, text="First Name").grid(row=0, column=0)
entry_first_name = tk.Entry(root)
entry_first_name.grid(row=0, column=1)

tk.Label(root, text="Last Name").grid(row=1, column=0)
entry_last_name = tk.Entry(root)
entry_last_name.grid(row=1, column=1)

tk.Label(root, text="Email").grid(row=2, column=0)
entry_email = tk.Entry(root)
entry_email.grid(row=2, column=1)

# Buttons for actions related to users
btn_add_user = tk.Button(root, text="Add User", command=add_user)
btn_add_user.grid(row=3, columnspan=2)

btn_send_emails = tk.Button(root, text="Send Daily Emails", command=send_emails)
btn_send_emails.grid(row=4, columnspan=2)

# Create input fields and labels for adding a message
tk.Label(root, text="Day of Year (1-365)").grid(row=5, column=0)
entry_date = tk.Entry(root)
entry_date.grid(row=5, column=1)

tk.Label(root, text="Message").grid(row=6, column=0)
entry_message = tk.Entry(root)
entry_message.grid(row=6, column=1)

tk.Label(root, text="Riddle").grid(row=7, column=0)
entry_riddle = tk.Entry(root)
entry_riddle.grid(row=7, column=1)

tk.Label(root, text="Yesterday's Riddle Answer").grid(row=8, column=0)
entry_yesterday_answer = tk.Entry(root)
entry_yesterday_answer.grid(row=8, column=1)

# Button for adding a message
btn_add_message = tk.Button(root, text="Add Message", command=add_message)
btn_add_message.grid(row=9, columnspan=2)

# Run the Tkinter event loop
root.mainloop()