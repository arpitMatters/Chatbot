from flask import *
import speech_recognition as sr
import webbrowser
import wolframalpha
import wikipedia
import time
import os
#import pyvona
import pyperclip
import win32com.client as wincl
import pythoncom
from flup.server.fcgi import WSGIServer
import mysql.connector

db_connection=mysql.connector.connect(host="localhost",user="root",passwd="RootUser@123",database="webassistant")
db_cursor=db_connection.cursor()

app=Flask(__name__)

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/newsignin')
def newsignin():
    return render_template("signin.html")

@app.route('/newlogin')
def newlogin():
    return render_template("login.html")

@app.route('/error')
def error():
    return render_template('error.html');

@app.route('/error2')
def error2():
    return render_template('error2.html');


@app.route('/signin',methods=['GET','POST'])
def signin():
    if request.method=='POST':
        username=request.form.get('user')
        emailid=request.form.get('email')
        password=request.form.get('pass')

        user_query="insert into userdetails(username,emailid,password) values(%s,%s,%s)"
        db_cursor.execute(user_query,(username,emailid,password))
        db_connection.commit()
        return render_template('login.html')

@app.route('/success')
def success():
    user=request.cookies.get('user')
    return render_template('success.html',name=10,user=user)

@app.route('/login',methods=['GET','POST'])
def login():
    if request.method=='POST':
        username=request.form.get('user')
        password=request.form.get('pass')
        user_query="select count(username) from userdetails where username=%s and password=%s"
        db_cursor.execute(user_query,(username,password))
        myresult=db_cursor.fetchone()
        if(myresult[0]!=0):
            resp=redirect(url_for('success'))
            resp.set_cookie('user',username)
            return resp
        else :
            return redirect(url_for('error'))


@app.route('/changepass')
def changepass():
    return render_template("changepass.html")


@app.route('/change',methods=['GET','POST'])
def change():
    if request.method=='POST':
        username=request.form.get('user')
        email=request.form.get('email')
        user_query="select count(username) from userdetails where username=%s and emailid=%s"
        db_cursor.execute(user_query,(username,email))
        myresult=db_cursor.fetchone()
        password=request.form.get('newpass')
        if(myresult[0]!=0):
            user_query="update userdetails set password=%s where username=%s and emailid=%s"
            db_cursor.execute(user_query,(password,username,email))
            db_connection.commit()
            return render_template("login.html")
        else:
            return redirect(url_for('error2'))


@app.route('/backtologin')
def back():
    return render_template('login.html')


@app.route('/useagain')
def useagain():
    user=request.cookies.get('user')
    return render_template('success.html',name=10,user=user)

@app.route('/launched')
def launched():

    pythoncom.CoInitialize()


    v =wincl.Dispatch("SAPI.SpVoice")                                         #api for win32com
    #v.voice_name = 'Emma'                                                                       #selecting ivona voice
    cl = wolframalpha.Client('JRGPQ3-G2GUPJ7UGU')                                                      #api for wolfram alpha
    att = cl.query('Test/Attempt')
    r = sr.Recognizer()                                                                         #starting the speech_recognition recognizer
    r.pause_threshold = 0.7                                                                     #it works with 1.2 as well
    r.energy_threshold = 400

    shell = wincl.Dispatch("WScript.Shell")                                           #to handle keyboard events
    v.Speak('Hello! For a list of commands, plese say "keyword list"...')
    print("For a list of commands, please say: 'keyword list'...")

    #List of Available Commands

    keywd = 'keyword list'
    google = 'search for'
    acad = 'academic search'
    sc = 'deep search'
    wkp = 'wiki page for'
    rdds = 'read this text'
    sav = 'save this text'
    bkmk = 'bookmark this page'
    vid = 'video for'
    wtis = 'what is'
    wtar = 'what are'
    whis = 'who is'
    whws = 'who was'
    when = 'when'
    where = 'where'
    how = 'how'
    paint = 'open paint'
    lsp = 'silence please'
    lsc = 'resume listening'
    stoplst = 'stop listening'

    while True:                                                                                 #The main loop

        with sr.Microphone() as source:

            try:
                print("Please Speak")
                audio = r.listen(source, timeout = None)                                        #instantiating the Microphone, (timeout = None) can be an option
                message = str(r.recognize_google(audio))
                print('You said: ' + message)

                if google in message:                                                           #what happens when google keyword is recognized

                    words = message.split()
                    del words[0:2]
                    st = ' '.join(words)
                    print('Google Results for: '+str(st))
                    url='http://google.com/search?q='+st
                    webbrowser.open(url)
                    v.Speak('Google Results for: '+str(st))

                elif acad in message:                                                           #what happens when acad keyword is recognized

                    words = message.split()
                    del words[0:2]
                    st = ' '.join(words)
                    print('Academic Results for: '+str(st))
                    url='https://scholar.google.ro/scholar?q='+st
                    webbrowser.open(url)
                    v.Speak('Academic Results for: '+str(st))

                elif wkp in message:                                                            #what happens when wkp keyword is recognized

                    try:

                        words = message.split()
                        del words[0:3]
                        st = ' '.join(words)
                        wkpres = wikipedia.summary(st, sentences=2)

                        try:

                            print('\n' + str(wkpres) + '\n')
                            v.Speak(wkpres)

                        except UnicodeEncodeError:
                            v.Speak(wkpres)

                    except wikipedia.exceptions.DisambiguationError as e:
                        print (e.options)
                        v.Speak("Too many results for this keyword. Please be more specific and try again")
                        continue

                    except wikipedia.exceptions.PageError as e:
                        print('The page does not exist')
                        v.Speak('The page does not exist')
                        continue

                elif sc in message:                                                             #what happens when sc keyword is recognized

                    try:
                        words = message.split()
                        del words[0:1]
                        st = ' '.join(words)
                        scq = cl.query(st)
                        sca = next(scq.results).text
                        print('The answer is: '+str(sca))
                        #url='http://www.wolframalpha.com/input/?i='+st
                        #webbrowser.open(url)
                        v.Speak('The answer is: '+str(sca))

                    except StopIteration:
                        print('Your question is ambiguous. Please try again!')
                        v.Speak('Your question is ambiguous. Please try again!')

                    else:
                        print('No query provided')

                elif paint in message:                                                          #what happens when paint keyword is recognized
                    #os.startfile('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Accessories\Paint')
                    os.system('mspaint')

                elif rdds in message:                                                           #what happens when rdds keyword is recognized
                    print("Reading your text")
                    words = message.split()
                    del words[0:3]
                    st = ' '.join(words)
                    pyperclip.copy(st)
                    v.Speak(pyperclip.paste())

                elif sav in message:                                                            #what happens when sav keyword is recognized
                    with open("D:/major/a.txt", 'a') as f:
                        words = message.split()
                        del words[0:3]
                        st = ' '.join(words)
                        pyperclip.copy(st)
                        f.write(pyperclip.paste())
                        print(pyperclip.paste())
                    print("Saving your text to file")
                    v.Speak("Saving your text to file")

                elif bkmk in message:                                                           #what happens when bkmk keyword is recognized
                    shell.SendKeys("^d")
                    v.Speak("Page bookmarked")

                elif keywd in message:                                                          #what happens when keywd keyword is recognized

                    print('')
                    print('Say ' + google + ' to return a Google search')
                    print('Say ' + acad+ ' to return a Google Scholar search')
                    print('Say ' + sc + ' to return a Wolfram Alpha query')
                    print('Say ' + wkp + ' to return a Wikipedia page')
                    #print('Say ' + book + ' to return an Amazon book search')
                    print('Say ' + rdds + ' to read the text you have highlighted and Ctrl+C (copied to clipboard)')
                    print('Say ' + sav + ' to save the text you have highlighted and Ctrl+C-ed (copied to clipboard) to a file')
                    print('Say ' + bkmk + ' to bookmark the page your are currently reading in your browser')
                    print('Say ' + vid + ' to return video results for your query')
                    print('For more general questions, ask them naturally and I will do my best to find a good answer')
                    print('Say ' + stoplst + ' to shut down')
                    print('')

                elif vid in message:                                                            #what happens when vid keyword is recognized

                    words = message.split()
                    del words[0:2]
                    st = ' '.join(words)
                    print('Video Results for: '+str(st))
                    url='https://www.youtube.com/results?search_query='+st
                    webbrowser.open(url)
                    v.Speak('Video Results for: '+str(st))

                elif wtis in message:                                                           #what happens when wtis keyword is recognized

                    try:

                        scq = cl.query(message)
                        sca = next(scq.results).text
                        print('The answer is: '+str(sca))
                        #url='http://www.wolframalpha.com/input/?i='+st
                        #webbrowser.open(url)
                        v.Speak('The answer is: '+str(sca))

                    except UnicodeEncodeError:

                        v.Speak('The answer is: '+str(sca))

                    except StopIteration:

                        words = message.split()
                        del words[0:2]
                        st = ' '.join(words)
                        print('Google Results for: '+str(st))
                        url='http://google.com/search?q='+st
                        webbrowser.open(url)
                        v.Speak('Google Results for: '+str(st))

                elif wtar in message:                                                           #what happens when wtar keyword is recognized

                    try:

                        scq = cl.query(message)
                        sca = next(scq.results).text
                        print('The answer is: '+str(sca))
                        #url='http://www.wolframalpha.com/input/?i='+st
                        #webbrowser.open(url)
                        v.Speak('The answer is: '+str(sca))

                    except UnicodeEncodeError:

                        v.Speak('The answer is: '+str(sca))

                    except StopIteration:

                        words = message.split()
                        del words[0:2]
                        st = ' '.join(words)
                        print('Google Results for: '+str(st))
                        url='http://google.com/search?q='+st
                        webbrowser.open(url)
                        v.Speak('Google Results for: '+str(st))

                elif whis in message:                                                           #what happens when whis keyword is recognized

                    try:

                        scq = cl.query(message)
                        sca = next(scq.results).text
                        print('\nThe answer is: '+str(sca)+'\n')
                        v.Speak('The answer is: '+str(sca))

                    except StopIteration:

                        try:

                            words = message.split()
                            del words[0:2]
                            st = ' '.join(words)
                            wkpres = wikipedia.summary(st, sentences=2)
                            print('\n' + str(wkpres) + '\n')
                            v.Speak(wkpres)

                        except UnicodeEncodeError:

                            v.Speak(wkpres)

                        except:

                            words = message.split()
                            del words[0:2]
                            st = ' '.join(words)
                            print('Google Results (last exception) for: '+str(st))
                            url='http://google.com/search?q='+st
                            webbrowser.open(url)
                            v.Speak('Google Results for: '+str(st))

                elif whws in message:                                                           #what happens when whws keyword is recognized

                    try:

                        scq = cl.query(message)
                        sca = next(scq.results).text
                        print('\nThe answer is: '+str(sca)+'\n')
                        v.Speak('The answer is: '+str(sca))

                    except StopIteration:

                        try:

                            words = message.split()
                            del words[0:2]
                            st = ' '.join(words)
                            wkpres = wikipedia.summary(st, sentences=2)
                            print('\n' + str(wkpres) + '\n')
                            v.Speak(wkpres)

                        except UnicodeEncodeError:

                            v.Speak(wkpres)

                        except:

                            words = message.split()
                            del words[0:2]
                            st = ' '.join(words)
                            print('Google Results for: '+str(st))
                            url='http://google.com/search?q='+st
                            webbrowser.open(url)
                            v.Speak('Google Results for: '+str(st))

                elif when in message:                                                         #what happens when 'when' keyword is recognized

                    try:

                        scq = cl.query(message)
                        sca = next(scq.results).text
                        print('\nThe answer is: '+str(sca)+'\n')
                        v.Speak('The answer is: '+str(sca))

                    except UnicodeEncodeError:

                        v.Speak('The answer is: '+str(sca))

                    except:

                        print('Google Results for: '+str(message))
                        url='http://google.com/search?q='+str(message)
                        webbrowser.open(url)
                        v.Speak('Google Results for: '+str(message))

                elif where in message:                                                        #what happens when 'where' keyword is recognized

                    try:

                        scq = cl.query(message)
                        sca = next(scq.results).text
                        print('\nThe answer is: '+str(sca)+'\n')
                        v.Speak('The answer is: '+str(sca))

                    except UnicodeEncodeError:

                        v.Speak('The answer is: '+str(sca))

                    except:

                        print('Google Results for: '+str(message))
                        url='http://google.com/search?q='+str(message)
                        webbrowser.open(url)
                        v.Speak('Google Results for: '+str(message))

                elif how in message:                                                          #what happens when 'how' keyword is recognized

                    try:

                        scq = cl.query(message)
                        sca = next(scq.results).text
                        print('\nThe answer is: '+str(sca)+'\n')
                        v.Speak('The answer is: '+str(sca))

                    except UnicodeEncodeError:

                        v.Speak('The answer is: '+str(sca))

                    except:

                        print('Google Results for: '+str(message))
                        url='http://google.com/search?q='+str(message)
                        webbrowser.open(url)
                        v.Speak('Google Results for: '+str(message))

                elif stoplst in message:                                                        #what happens when stoplst keyword is recognized
                    v.Speak("I am shutting down")
                    print("Shutting down...")
                    break

                elif lsp in message:

                    v.Speak('Listening is paused')
                    print('Listening is paused')
                    r2 = sr.Recognizer()
                    r2.pause_threshold = 0.7
                    r2.energy_threshold = 400

                    while True:

                        with sr.Microphone() as source2:

                            try:

                                audio2 = r2.listen(source2, timeout = None)
                                message2 = str(r.recognize_google(audio2))

                                if lsc in message2:
                                    v.Speak('I am listening')
                                    break

                                else:
                                    continue

                            except sr.UnknownValueError:
                                print("Listening is paused. Say resume listening when you're ready...")

                            except sr.RequestError:
                                v.Speak("I'm sorry, I couldn't reach google")
                                print("I'm sorry, I couldn't reach google")


                else:
                    pass

            except sr.UnknownValueError:
                print("For a list of commands, say: 'keyword list'...")

            except sr.RequestError:
                v.Speak("I'm sorry, I couldn't reach google")
                print("I'm sorry, I couldn't reach google")

        time.sleep(0.3)
    return redirect(url_for('success'))

if __name__=='__main__':
    app.run(debug=True)
