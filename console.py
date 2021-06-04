# General imports
from datetime import datetime as dt
from datetime import timedelta as timedelta

from docbinder import DocBinder


print("imported datetime as dt, timedelta")

def diff(a, b):
    print("( {} - {} ) / {} = ".format(b,a,a))
    return (b-a)/a

print("defined diff(a,b)")

def daysfromtoday(days):
    future_day = dt.now() + timedelta(days=days)
    print(future_day.strftime("%a %m/%d/%y"))

db = DocBinder()
