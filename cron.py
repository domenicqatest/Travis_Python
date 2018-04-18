### Place cron.py in the same folder as the pytests. ###

import schedule
import time
import os
import datetime

def job():
    print('\n')  # adds line break
    print("The date / time is:")
    print datetime.datetime.now()

    # pytests (or test entire folder)
    os.system("pytest amp_fails_only.py -s")

#schedule.every(10).seconds.do(job)
#schedule.every(2).minutes.do(job)
#schedule.every(2).hours.do(job)
#schedule.every().day.at("7:00").do(job)
#schedule.every().monday.do(job)
schedule.every().monday.at("7:00").do(job)
schedule.every().wednesday.at("7:00").do(job)
schedule.every().friday.at("7:00").do(job)
#schedule.every(14).days.at("7:00").do(job)


while True:
    schedule.run_pending()
    time.sleep(1)