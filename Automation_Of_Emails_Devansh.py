from email.message import EmailMessage
import os
import ssl
import smtplib
import win32com.client
import speech_recognition as sr
import webbrowser
def speak(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            print("Recognizing...")
            body = r.recognize_google(audio, language="en-in")
            print(f"User said: {body}")
            return body
        except Exception as e:
            return "Some Error Occurred, Sorry"
if __name__ == '__main__':
    print("Hello, I am Devansh's A.I Robot , I can help you with automation of Emails.")
    speak("Hello, I am Devansh's A.I Robot , I can help you with automation of Emails.")
    email_sender = "sender's email id"
    email_password = "sender's password"
    email_receiver = "receiver's email id"
    subject = input("Enter Subject:")
    print("Press 1 to type the body of your email.")
    print("Press 2 to speak the body of your email.")
    choice = input("Your choice: ")
    if choice == "1":
        body = input("Enter the body of your email: ")
    elif choice == "2":
        print("Listening...")
        body = takeCommand()
    else:
        print("Invalid choice. Exiting...")
        exit()
    em = EmailMessage()
    em['From'] = email_sender
    em['To'] = email_receiver
    em['Subject'] = subject
    em.set_content(body)

    # attachment_path = r'path_of_your_file' #Don't comment this portion if you want to send attachments also
    # with open(attachment_path, 'rb') as attachment:
    #     em.add_attachment(
    #         attachment.read(),
    #         maintype='application',
    #         subtype='octet-stream',
    #         filename='file.txt' #file.txt is the name of the file
    #     )

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(email_sender, email_password)
        smtp.sendmail(email_sender, email_receiver, em.as_string())
    print("Email Successfully Sent")
    speak("Email Successfully Sent")
