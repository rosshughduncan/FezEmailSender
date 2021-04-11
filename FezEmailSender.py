import os
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Email, To, Content
from datetime import datetime
import json

def convertToInt(msg):
    flag = False
    while flag == False:
        try:
            result = int(input(msg))
            flag = True
        except:
            print("Error! Check your number format and try again")
    return result

def inputNoBlank(msg):
    flag = False
    while flag == False:
        result = input(msg)
        if len(result) > 0:
            flag = True
        else:
            print("Empty input! Please try again")
    return result

# Lists
emailList = []
paragraphs = []

# Create an emails directory if there isn't one already
try:
    os.mkdir('./Emails')
except:
    print('===== The directory Emails already exists. Emails will be stored here in .json format =====\n')
else:
    print('===== The directory Emails has been created in the folder where this script is located ====')
    print('===== It contains all of the emails which you have sent in .json format =====\n')

# Enter the sender details and pass them to the helpers
fromEmail = Email(inputNoBlank("Enter the email address to send from: "))
emailSubject = inputNoBlank("Enter the email subject for all of the recipients: ")
yourName = inputNoBlank("Enter your full name: ")

# Enter list of email addresses
flag = True
while flag == True:
    emailList.append({'address': inputNoBlank("Enter a recipient email address: "), 'firstName': inputNoBlank("Enter a first name: "), 'surname': inputNoBlank("Enter a surname: ")})
    check = input("Stop adding? (type y to end, or any other key to add another address)")
    if check == 'y':
        flag = False

# Construct message
numberParagraphs = convertToInt("Enter the number of paragraphs to use: ")
for currentIndex in range(0, numberParagraphs):
    paragraphs.append(inputNoBlank(f"Enter paragraph {currentIndex + 1}:\n"))
lastSentence = inputNoBlank("Enter your last sentence:\n")
endGreeting = inputNoBlank("Enter your end greeting:\n")

# Send email for each recipient
for currentPerson in emailList:
    stillWriting = True
    while stillWriting == True:

        # Writing message
        print(f"\n===== Now emailing {currentPerson['firstName']} {currentPerson['surname']} =====\n")
        messageText = f"<p>Hello {currentPerson['firstName']}</p>"
        for currentParagraph in paragraphs:
            print(currentParagraph)
            messageText = messageText + "<p>" + currentParagraph + " " + inputNoBlank("Continue this paragraph:\n") + "</p>"
            print()
        messageText = messageText + '<p>' + lastSentence + '</p><p>' + endGreeting + "</p><p>" + yourName + '</p>'
        print(f"\nFull message for {currentPerson['firstName']} {currentPerson['surname']}:\n\n***")
        messageTextToDisplay = messageText.replace('</p>', '</p>\n\n')
        print(messageTextToDisplay)
        print('***\n')
        print(f"Recipient email address: {currentPerson['address']}")
        check = input("\nAre you sure you want to send this message? This action cannot be undone (type y)")

        # Confirm from the user that they want to send the message
        if check == 'y':
            stillWriting = False

            # Pass the to address to the helper
            toEmail = To(currentPerson['address'])
            content = Content('text/html', messageText)

            # Construct the message
            message = Mail(fromEmail, toEmail, emailSubject, content)

            # Save a .json representation of the message
            mailJSON = message.get()

            # Add sending and receiving names to the message
            mailJSON['from']['name'] = yourName
            mailJSON['personalizations'][0]['to'][0]['name'] = currentPerson['firstName'] + " " + currentPerson['surname']

            # Save .json file to disk
            timestamp = datetime.now()
            fileName = f"{timestamp}".replace(" ", "_") + "_" + currentPerson['firstName'] + "-" + currentPerson['surname']
            fileName.replace('/', '-')
            fileName.replace(' ', '-')
            with open(f"./Emails/{fileName}.json", 'w') as outfile:
                json.dump(mailJSON, outfile)

            # Try to send the message
            responseFlag = False
            while responseFlag == False:
                try:
                    # Use SendGrid API linked to the OS
                    sg = SendGridAPIClient(api_key=os.environ.get('SENDGRID_API_KEY'))
                    response = sg.client.mail.send.post(request_body=mailJSON)
                    print(f"\nSendGrid response for {currentPerson['address']}")
                    print(response.status_code)
                    print(response.body)
                    print(response.headers)
                    responseFlag = True
                except Exception as e:
                    print(e)
                    retryCheck = input("\nTry again? Type y to try again. Press any other key to move onto the next email (note this faulty email down)")
                    if retryCheck != "y":
                        responseFlag = True
        
        # If user does not like the message, they can enter it again
        else:
            print("===== Enter the response for your message again below =====\n")

print("\n==== Operation complete =====")