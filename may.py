from win32com.client import Dispatch
from pygame import mixer
mixer.init()
from playsound import playsound
import speech_recognition as sprec
import time
import youtube_dl
import wikipedia
import webbrowser
import pafy
import os
import myweathercalls
import requests
import json
import re
import random
import threading
from gtts import gTTS


def speakmay1(talkie):
	spk = Dispatch("SAPI.SpVoice")
	spk.Speak(f'. . .{talkie}')

def speakmay2(talkie):
	mytalks = str(talkie)
	language = 'hi'
	myobj = gTTS(text=mytalks,lang=language,tld='co.in',slow=False)
	date_string = time.strftime("%d%m%Y%H%M%S",time.localtime())
	filename = "voice"+date_string+".mp3"
	myobj.save(filename)
	playsound(filename)
	os.remove(filename)
	# mixer.music.load("talkie.mp3")
	# mixer.music.play()
	# while mixer.music.get_busy():
	# 	pass
	# mixer.music.load("talkie_piggy.mp3")
	# os.remove("talkie.mp3")

def speakmay(talkie):
	try:
		speakmay1(talkie)
	except Exception as e:
		speakmay2(talkie)

def wishmemay():
	nowhr = int(time.strftime("%H",time.localtime()))

	if nowhr < 12:
		speakmay("Good Morning! It's me, May. Have a good day.")
	elif nowhr >= 12 and nowhr <= 16:
		speakmay("Good Afternoon!")
	elif nowhr > 16 and nowhr <= 21 :
		speakmay("Good Evening!")
	else:
		speakmay("Good Night! Take Care & lots of love.")
	speakmay("How can I help you!?")

def listen():
	r = sprec.Recognizer()
	with sprec.Microphone() as source:
		print("\nSay.<3")
		r.pause_threshold = 1
		audio = r.listen(source,timeout=7,phrase_time_limit=12)
		# audio = r.listen(source)

	try:
		print("Recognising...")
		usertalks = r.recognize_google(audio, language="en-in")
		print(f'\n ~ {usertalks}\n')
		return usertalks
	except Exception as e:
		if breakingnum >= 10 and mixer.music.get_busy()==0 :
			print("\nDon't talk to me. Bye.")
			speakmay("Don't talk to me. Bye.")
		else:
			print("Say that again please.")
			time.sleep(1)
		return None #not working as of now

def ytfirsturlreturn(query):

	global queuedict

	results = requests.get(f'https://www.youtube.com/results?search_query={query.replace(" ","+")}').text

	resultsIndex = results.index('{"videoId":')

	resultsRefined = results[resultsIndex:resultsIndex+3000]

	found = re.search(r'{"videoId":"[-.\d\w]+', resultsRefined)[0].split("\"")[3]

	foundindex = resultsRefined.index(found)

	containingtitle = resultsRefined[foundindex:foundindex+1500]

	title = re.findall(r'{"text":"[^"]+"}',containingtitle)[0].split('"')[3]

	url = f'https://youtu.be/{found}'

	queuedict[url]=title

	return url

def musicplay(loop):
	musicdir = [ item for item in os.listdir(r'F:/Anik/My Musics/English musics') if os.path.splitext(item)[1] == ".mp3"]
	for index,item in enumerate(musicdir,1):
		print(f'{index}. {item}')
	musicwant = int(input("\nWhich song you want to play!?\n"))
	os.chdir(r'F:/Anik/My Musics/English musics')
	# mixer.set_num_channels(1)
	mixer.music.load(musicdir[musicwant-1])
	mixer.music.fadeout(10000)
	mixer.music.play(loop)

def musicqueue():
	musicdir = [ item for item in os.listdir(r'F:/Anik/My Musics/English musics') if os.path.splitext(item)[1] == ".mp3"]
	for index,item in enumerate(musicdir,1):
		print(f'{index}. {item}')
	musicwant = int(input("\nWhich song you want to add to queue!?\n"))
	os.chdir(r'F:/Anik/My Musics/English musics')
	print(f'Playing {musicdir[musicwant-1].split(".")[0]}')
	mixer.music.queue(musicdir[musicwant-1])

def musicrandomplay(loop):
	musicdir = [ item for item in os.listdir(r'F:/Anik/My Musics/English musics') if os.path.splitext(item)[1] == ".mp3"]
	os.chdir(r'F:/Anik/My Musics/English musics')
	# mixer.set_num_channels(1)
	musicwant = random.randint(0,len(musicdir)-1)
	print(f'Playing {musicdir[musicwant].split(".")[0]}')
	speakmay(f'Playing {musicdir[musicwant].split(".")[0]}')
	mixer.music.load(musicdir[musicwant])
	# mixer.music.fadeout(10000)
	mixer.music.play(loop)

def ytdlMusicPlay(keyword):
	infoFile = ytdl.extract_info(keyword,download=True)

	nowplaying = infoFile['entries'][0]['title']
	nowid = infoFile['entries'][0]['id']

	if mixer.music.get_busy()==0:
		print(f"Playing {nowplaying} !")

		mixer.music.load(fr'{nowid}.mp3')
		mixer.music.play()

	else:
		# queue.append(nowid)
		mixer.music.queue(fr'{nowid}.mp3')

		print("\nWait for this song to play to add another song to the queue!")
		time.sleep(2)


def newsreturn(numberofnews,topic):
	newsap = requests.get(f"https://newsapi.org/v2/top-headlines?country=in&pageSize={numberofnews+1}&page=1category={topic}&apiKey=ebfb37ba82c542b1979be299547484dd").text
	jsondata = json.loads(newsap)

	newslist = [ jsondata['articles'][i]['title'] for i in range(numberofnews) ]
	return newslist

def quotereturn():
	url = "https://zenquotes.io/?api=random"
	quotetext = requests.get(url).text
	return [json.loads(quotetext)[0]['q'],json.loads(quotetext)[0]['a']]

def myalarm(hrs,minutes,ampm):
	# consent12hr = print("\nAt what time you want to be reminded?")
	# hrs = int(input("Enter the Hours(12HR format): "))
	# minutes = int(input("Enter the Minutes: "))
	# ampm = input("Enter AM/PM: ")

	# nowhr = int(time.strftime("%I",time.localtime()))
	# nowmin = int(time.strftime("%M",time.localtime()))
	# nowampm = time.strftime("%p",time.localtime()).lower()

	# yourtime = f'{hrs.lower()}:{minutes.lower()}{ampm.lower()}'
	while True:
		# if time.strftime("%I:%M%p",time.localtime()).lower() == yourtime :
		if int(time.strftime("%I",time.localtime())) == hrs and int(time.strftime("%M",time.localtime())) == minutes and time.strftime("%p",time.localtime()).lower() == ampm :
			speakmay(f"Your Alarm is Up.")
			print(f"Your Alarm is Up.")
			# musicrandomplay(0)
			# time.sleep(30)
			# mixer.music.stop()
			break



ytdl_format_options = {
    # 'format': 'bestaudio/best',
    'format': 'worstaudio/worst',
    'outtmpl': '%(id)s.%(ext)s',

    'postprocessors':[
    {'key':'FFmpegExtractAudio'},
    {'key':'FFmpegExtractAudio',
    'preferredcodec':'mp3'}],

    'restrictfilenames': True,
    'noplaylist': True,
    # 'forceduration': True,
    'nocheckcertificate': True,
    'ignoreerrors': False,
    'logtostderr': False,
    'quiet': True,
    'no_warnings': True,
    'default_search': 'auto',
    'source_address': '0.0.0.0' # bind to ipv4 since ipv6 addresses cause issues sometimes
}

ytdl = youtube_dl.YoutubeDL(ytdl_format_options)




if __name__ == '__main__':
	os.system('cls')

	wishmemay()

	queuedict={}

	while True:
		breakingnum = 0
		while True:
			os.system('cls')
			try:
				say = listen()
				if say is not None and breakingnum < 10:
					breakingnum = 0
					break
				elif mixer.music.get_busy()==1:
					breakingnum = 1
				elif breakingnum >= 10  and mixer.music.get_busy()==0 :
					exit()
				else:
					breakingnum += 1
			except Exception as e:
				speakmay("Speak a bit earlier!")

		# wants = input("\n\nEnter: ").lower()
		wants = say.lower()
		# listen()
		if mixer.music.get_busy()==1:
			if "pause" in wants:
				mixer.music.pause()
				continue

			elif "stop" in wants:
				mixer.music.stop()
				continue

			elif "rewind" in wants:
				mixer.music.rewind()
				continue

			elif "volume" in wants:
				try:
					vol = re.findall(r'\d+',wants)[0]
					if "percent" in wants or "%" in wants:
						mixer.music.set_volume(int(vol)/100)
					else:
						mixer.music.set_volume(int(vol)/10)
				except Exception as e:
					print("Please give a value of volume.")
					speakmay("Please give a value of volume.")
				finally:
					continue
			elif "next" in wants:
				musicrandomplay(0)
				continue

		if wants == "exit" or 'bye' in wants:
			speakmay("Bye Bye! Take Care.")
			break

		elif "sleep" in wants and [x for x in wants if x.isnumeric()]!=[]:
			if "min" in wants:
				time.sleep(int(re.findall(r'\d+',wants)[0])*60)
			else:
				time.sleep(int(re.findall(r'\d+',wants)[0]))


		elif "wikipedia" in wants:
			try:
				speakmay("Found this from Wikipedia.")
				# find = wants.split(" ")[0]
				results = wikipedia.summary(wants.replace("wikipedia",""),sentences=2)
				print(results)
				speakmay(results)

			except Exception as e:
				print("Couldn't fetch weather. Try again.")



		elif "music" in wants:
			if "play" in wants :
				if mixer.music.get_busy()==0 and "loop" not in wants and "queue" not in wants and "random" not in wants:
					musicplay(0)
				elif mixer.music.get_busy()==0 and "loop" in wants and "queue" not in wants and "random" not in wants:
					musicplay(-1)
				elif mixer.music.get_busy()==0 and "loop" not in wants and "queue" not in wants and "random" in wants:
					musicrandomplay(0)
				elif mixer.music.get_busy()==0 and "loop" in wants and "queue" not in wants and "random" in wants:
					musicrandomplay(-1)
				else:
					if "queue" in wants and "loop" not in wants:
						musicqueue()
					elif "loop" in wants:
						mixer.music.unload()
						musicplay(-1)
					else:
						mixer.music.unload()
						musicplay(0)

			elif "resume" in wants:
				mixer.music.unpause()

		elif "time" in wants:
			timenow = f'Time is {time.strftime("%I:%M%p",time.localtime())}.'
			print(timenow)
			speakmay(timenow)

		elif "weather" in wants:
			try:
				speakmay(myweathercalls.OpenweatherAPIcalltext())
			except Exception as e:
				print("Couldn't fetch weather. Try again.")

		elif "news" in wants:
			try:
				# business entertainment general health science sports technology
				if "business" in wants:
					topic = "business"
				elif "entertainment" in wants:
					topic = "entertainment"
				elif "health" in wants:
					topic = "health"
				elif "science" in wants:
					topic = "science"
				elif "sport" in wants:
					topic = "sports"
				elif "tech" in wants:
					topic = "technology"
				else:
					topic = "general"
					
				numnews = re.findall(r'\d+',wants) 
				if numnews == ["0"] or numnews == [] :
					numberofnews = 5
				else:
					numberofnews = int(numnews[0])
				print(f"\nTop {topic.title()} Headlines of Today:\n")
				for index,item in enumerate(newsreturn(numberofnews,topic),1):
					print(f'{index}. {item}')
					speakmay(item)

			except Exception as e:
				print("Couldn't get news. Try again.")
			
		elif "date" in wants or "today" in wants:
			todayday = time.strftime("%A",time.localtime())
			todaydate = time.strftime("%d %B",time.localtime())
			daydatewish = f'Today is {todayday} and Date is {todaydate}.'
			print(daydatewish)
			speakmay(daydatewish)

		elif "open" in wants: # and "sublime" in wants:
			# findd = re.compile(r'')
			# openwhat = re.findall(r'open \S+ \S',str2)[0].split(" ")

			if "google" in wants:
				webbrowser.open("www.google.com")

			elif "youtube" in wants:
				webbrowser.open("www.youtube.com")

			elif "whatsapp" in wants:
				webbrowser.open("https://web.whatsapp.com")

			elif "chrome" in wants:
				os.startfile("C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe")

			elif "sublime" in wants:
				os.startfile(r'C:\Program Files\Sublime Text 3\sublime_text.exe')


		elif "play" in wants and "youtube" in wants:
			url = ytfirsturlreturn(wants.replace("play","").replace("youtube","").replace("on",""))
			print(f"Playing {queuedict[url]} on Youtube")
			speakmay(f"Playing {queuedict[url]} on Youtube")
			webbrowser.open_new_tab(url)

		elif "play" in wants:
			keyword = wants.replace("play","")

			print("Please Wait...")

			ytdlMusicPlay(keyword)

		elif "search" in wants and "google" in wants:
			url = f'https://www.google.com/search?q={wants.replace("search ","").replace("google ","").replace(" ","+")}'
			webbrowser.open_new_tab(url)

		elif "search" in wants and "youtube" in wants:
			keyword = wants.replace("search ","").replace("youtube ","")
			print(f"Searching for {keyword}")
			url = f'https://www.youtube.com/results?search_query={keyword.replace(" ","+")}'
			webbrowser.open_new_tab(url)

		elif "search" in wants:
			try:
				keyword = wants.replace("search","")
				print(f"Searching for {keyword}")
				results = wikipedia.summary(keyword,sentences=2)
				if "==" in results:
					print(wikipedia.summary(keyword))
					speakmay(wikipedia.summary(keyword))
				else:
					print(results)
					speakmay(results)
			except Exception as e:
				print(e)

		elif "joke" in wants:
			try:
				url = "https://api.yomomma.info/"
				joketext = requests.get(url)
				jokejson = json.loads(joketext.text)
				speakmay(jokejson['joke'])
			except Exception as e:
				print("Couldn't get joke. Cause your life's a joke.")

		elif "insult me" in wants:
			try:
				url = "https://evilinsult.com/generate_insult.php"
				insulttext = requests.get(url).text
				speakmay(insulttext)
			except Exception as e:
				print("Couldn't get Insult. Cause your life's a joke.")
		
		elif "quote" in wants:
			try:
				quote = quotereturn()[0]
				author = quotereturn()[1]
				print(quote,"  -", author)
				speakmay(quote)
			except Exception as e:
				print("Couldn't get quote. Try again.")

		elif "alarm" in wants or "reminder" in wants:
			consent12hr = print("\nAt what time you want to be reminded?")
			hrs = int(input("Enter the Hours(12HR format): "))
			minutes = int(input("Enter the Minutes: "))
			ampm = input("Enter AM/PM: ").lower()

			try:
				alarmthread = threading.Thread(target = myalarm, args=(hrs,minutes,ampm,))
				alarmthread.start()
			except Exception as e:
				print("Couldn't set alarm/reminder.")

		elif "say" in wants[0:4] or "speak" in wants[0:6]:
			speakmay(wants.replace("say","").replace("speak",""))

		else:
			speakmay("Sorry, I can't help with this.")

		time.sleep(2)



