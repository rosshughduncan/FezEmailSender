import os
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail

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

# Enter sender details
fromEmail = inputNoBlank("Enter the email address to send from: ")
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

        # Upon confirmation from the user, the message is sent
        if check == 'y':
            stillWriting = False
            message = Mail(
                from_email=fromEmail,
                to_emails=f"{currentPerson['address']}",
                subject=emailSubject,
                html_content=f"{messageText}"
            )
            responseFlag = False
            while responseFlag == False:
                try:

                    # Use SendGrid API linked to the OS
                    sg = SendGridAPIClient(os.environ.get('SENDGRID_API_KEY'))
                    response = sg.send(message)
                    print(f"\nSendGrid response for {currentPerson['address']}")
                    print(response.status_code)
                    print(response.body)
                    print(response.headers)
                    responseFlag = True
                except Exception as e:
                    print("Error!\n" + e)
                    retryCheck = input("\nTry again? Type y to try again. Press any other key to move onto the next email (note this faulty email down)")
                    if retryCheck != "y":
                        responseFlag = True
        
        # If user does not like the message, they can enter it again
        else:
            print("===== Enter the response for your message again below =====\n")

print("\n==== Operation complete =====")