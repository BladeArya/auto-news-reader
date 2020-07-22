import requests

def TopNews():
    url = ('http://newsapi.org/v2/top-headlines?country=in&apiKey=96623079d8744e49902c0d189577842a')
    response = requests.get(url).json()
    article=response["articles"]
    headline_results=[]
    for ar in article:
        headline_results.append(ar['title'])
    for i in range(len(headline_results)):
        print(i+1,headline_results[i])

    from win32com.client import Dispatch
    speak=Dispatch("SAPI.Spvoice")
    speak.Speak(headline_results)

if __name__=='__main__':
    TopNews()
