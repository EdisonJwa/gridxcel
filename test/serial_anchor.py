
import datetime, sys
def from_serial(s):
    base = datetime.datetime(1899,12,30)
    return base + datetime.timedelta(days=s)
for s in [1, 59, 61, 25569, 44927, 45292, 45293, 45293.5]:
    print(s, "->", from_serial(s).strftime("%Y-%m-%d %H:%M:%S"), "weekday", from_serial(s).strftime("%a"))
