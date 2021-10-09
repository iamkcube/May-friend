def myvoice(talkie):
	from win32com.client import Dispatch	
	spk = Dispatch("SAPI.SpVoice")

	spk.Speak(talkie)

def OpenweatherAPIcalltext():
	'''Uses Open Weather API to fetch weather data of Bhubaneswar'''
	import json
	import requests

	weatherjson = requests.get("https://api.openweathermap.org/data/2.5/weather?lat=20.292&lon=85.838&units=metric&appid=145c485b75dd3ef3084c53041a74560f")

	textparsed = json.loads(weatherjson.text)

	weather = textparsed['weather'][0]['description']
	temperature = textparsed['main']['temp']
	humidity = textparsed['main']['humidity']
	feels_like = textparsed['main']['feels_like']
	clouds = textparsed['clouds']['all']
	wind = textparsed['wind']['speed']
	name = textparsed['name']

	openweathercall = f'''The weather is {weather} in {name}. 
	Temperature is {temperature}°C. 
	Due to {humidity}% humidity, it feels like {feels_like}°C.
	Cloud Coverage is {clouds}%.
	And Wind Speed is {wind}m/s.'''

	return openweathercall



def FreeweatherAPIcalltext():
	'''Uses Free Weather API to fetch weather data of Bhubaneswar'''
	import json
	import requests

	weatherjson = requests.get("https://api.weatherapi.com/v1/current.json?key=dbe8bfdead3547a6bea133230213009&q=Bhubaneswar&days=7&alerts=yes")

	textparsed = json.loads(weatherjson.text)

	name = textparsed['location']['name']
	weather = textparsed['current']['condition']['text']
	temperature = textparsed['current']['temp_c']
	humidity = textparsed['current']['humidity']
	feels_like = textparsed['current']['feelslike_c']
	clouds = textparsed['current']['cloud']
	wind = textparsed['current']['wind_kph']
	isday = textparsed['current']['is_day']

	precip_mm = textparsed['current']['precip_mm']
	if precip_mm == 0:
		rain = "No Rain today. "
	else:
		rain = f"It rained {precip_mm} inches in {name}"


	if isday == 0:
		wish = "Good Evening, Kalinga."
	else:
		wish = "Good Morning, Kalinga."

	freeweathercall = f'''{wish} The weather is {weather} in {name}. 
	Temperature is {temperature}°C. 
	Due to {humidity}% humidity, it feels like {feels_like}°C.
	Cloud Coverage is {clouds}%. 
	{rain}
	And Wind Speed is {wind}km/h.'''

	return freeweathercall





if __name__ == '__main__':
	# print(weathercalltext())
	myvoice(OpenweatherAPIcalltext())
	print(OpenweatherAPIcalltext())
	# myvoice(FreeweatherAPIcalltext())