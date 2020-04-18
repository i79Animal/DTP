from datetime import datetime, timedelta

today = datetime.today()
monday = today - timedelta(days=7)
sunday = today
print(monday.strftime("%d.%m")+'-'+sunday.strftime("%d.%m"))