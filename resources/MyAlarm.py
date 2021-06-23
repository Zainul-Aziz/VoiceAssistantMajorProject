import datetime
import winsound


def alarm(Timing):
    alTime = str(datetime.datetime.now().strptime(Timing, "%I:%M %p"))
    alTime = alTime[11:-3]
    Horeal = alTime[:2]
    Horeal = int(Horeal)
    Mireal = alTime[3:5]
    Mireal = int(Mireal)
    print(f"Done, alarm is set for {Timing}")

    while True:
        if Horeal == datetime.datetime.now().hour:
            if Mireal == datetime.datetime.now().minute:
                print("ALarm is running")
                winsound.PlaySound('abc', winsound.SND_FILENAME)
            elif Mireal < datetime.datetime.now().minute:
                break

# if __name__ == '__main__':
#     alarm('9:16 AM')
