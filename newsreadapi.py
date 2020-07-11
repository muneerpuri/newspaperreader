import requests
import json
from win32com.client import Dispatch
req = requests.get("https://newsapi.org/v2/top-headlines?country=in&apiKey=0eec55177b584c788526f5e8db689471")
reqd = req.text
parsed = json.loads(reqd)
dic1 = parsed['articles']
print(type(dic1))
speak = Dispatch("SAPI.spvoice")
speak.speak(" Top Ten Headlines are")
for i in range(10):
    speakval = dic1[i]['title']
    speak = Dispatch("SAPI.spvoice")
    d=i
    speak.speak("news number " + str(d+1) + "is ." + speakval)
