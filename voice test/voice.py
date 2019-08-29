import pyttsx3
import datetime

mytime = datetime.datetime.now()

if int(mytime.hour) < 12:
    print('Good morning')

engine = pyttsx3.init()
engine.setProperty('rate',130)
# engine.say('The quick brown fox jumped over the lazy dog.')

# voices = engine.getProperty('voices')

# for voice in voices:
    # print("Voice:")
    # print(" - ID: %s" % voice.id)
    # print(" - Name: %s" % voice.name)
    # print(" - Languages: %s" % voice.languages)
    # print(" - Gender: %s" % voice.gender)
    # print(" - Age: %s" % voice.age)

#engine.say('There is, 1, email in inbox' )

#engine.say('Oops! There was an error, in script. I am, terminating.' )
#engine.say('New, email received' )
#engine.say('Received time, ' + '2019-Dec-14 14:20 PM' )

engine.say('Good Morning, Mihindhu' )
# engine.say('There are, 9, emails in inbox')
# engine.say('among them, you have...')
# engine.say('1, volvo emails')
# engine.say('3, sas emails')
# engine.say('4, husqvarna email')


ema = 'Support <sas.surveillance@tradetechconsulting.com>'
#ema = '<BradyCreditRiskFileMover@rg19.se> '

client = ema.split('@')
client2 = client[-1].split('.')[0]
client3 = ema.split('<')[1].split('>')[0]

print(client2)
# engine.say(client3)
engine.runAndWait()