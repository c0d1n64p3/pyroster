
import datetime as dt


class Calendar:
    def __init__(self,name):
        self.name = name
        self.event_list = []

    def add_event(self, event):
        self.event_list.append(event)

    def remove_event(self, event):
        pass

    def print(self):
        calendar = f"""BEGIN:VCALENDAR
PRODID:-//NORMAN@NORTAN-DS.DE
VERSION:2.0
CALSCALE:GREGORIAN
X-WR-CALNAME:{self.name}{merge(self.event_list)}
END:VCALENDAR"""
        return calendar


    def save_ics(self, filename='worktime.ics', path=""):
        with open(filename, 'w') as icsfile:
            icsfile.write(self.print())

class Event:
    def __init__(self, organizer, summary, description, begin, end, busy, all_day, country_code, location=None, attendees=None):

        self.organizer = organizer       # -> str
        self.summary = summary          # -> str
        self.description = description  # -> str
        self.begin = begin              # -> datetime
        self.end = end                  # -> datetime
        self.busy = busy                # -> bool
        self.all_day = all_day          # -> bool
        self.location = location        # -> str
        self.attendees = attendees      # -> list
        self.uid = self.gen_uid()
        self.freq_rule = None           # -> str
        self.alarm_list = []            # -> list

    def gen_uid(self):
        time_str = dt.datetime.now().strftime("%Y%m%dT%H%M%S-%f")
        uid = f"{time_str}@nortan-ds.de"
        return uid

    def get_timezone(self):
        return "Europe/Berlin"

    def set_feq_rule(self, freq, bydate=None, interval=None, end=None, wkst="MO"):
        # set frequency
        # freq -> y, m, w, d
        freq_list = ["YEARLY", "MONTHLY", "WEEKLY", "DAILY"]
        for i in freq_list:
            if i[0] == freq.upper():
                freq = i
        rule = f"FREQ={freq}"

        # set interval
        if interval:
            rule += f";INTERVAL={interval}"

        # set end
        if isinstance(end, int):
            rule += f";COUNT={end}"
        elif isinstance(end, dt.datetime):
            rule += f";UNTIL={format_time(end)}Z"

        # set weekstart
        rule += f";WKST={wkst}"

        # set days
        if isinstance(bydate, str):
            rule += f";BYDAY={bydate.upper()}"
        elif isinstance(bydate, int):
            rule += f";BYMONTH={bydate}"

        self.freq_rule = rule

    def remove_freq_rule(self):
        self.freq_rule = None

    def add_alarm(self, trigger, description, action="d"):
        alarm = Alarm(self.begin, trigger, description, action)
        self.alarm_list.append(alarm)

    def remove_alarm(self,description=None, listindex=None):
        if description:
            for alarm in self.alarm_list:
                if alarm.description == description:
                    self.alarm_list.remove(alarm)
                    return
        elif listindex:
            del self.alarm_list[listindex]

    def print(self):

        if self.freq_rule == None:
            rule = ""

        else:
            print("rrule true")
            rule = f"\nRRULE:{self.freq_rule}"


        if self.busy:
            transp = "OPAQUE"
        else:
            transp = "TRANSPARENT"

        if self.all_day:
            if isinstance(self.begin, dt.datetime):
                self.begin = self.begin.date()
                self.end = self.end.date()

        if isinstance(self.begin, dt.date):
            complement = "VALUE=DATE"
        else:
            complement = f"TZID={self.get_timezone()}"
        now = dt.datetime.now()

        event = f"""
BEGIN:VEVENT
SUMMARY:{self.summary}
ORGANIZER:CN={self.organizer}
DTSTART;{complement}:{format_time(self.begin)}
DTEND;{complement}:{format_time(self.end)}
DTSTAMP:{format_time(now)}Z
UID:{self.uid}
SEQUENCE:0
CLASS:PRIVATE
CREATED:{format_time(now)}Z
DESCRIPTION:{self.description}
LAST-MODIFIED:{format_time(now)}Z
LOCATION:{self.location}{rule}
STATUS:CONFIRMED
TRANSP:{transp}{merge(self.alarm_list)}
END:VEVENT"""

        return event

class Alarm:
    def __init__(self, begin, trigger, description, action):
        self.description = description  # -> str
        self.trigger = trigger          # -> timedelta or datetime
        self.action = action            # -> str
        self.days = None                # -> int
        self.hours = None               # -> int
        self.minutes = None             # -> int
        self.seconds = None             # -> int

        if isinstance(self.trigger, dt.datetime):
            self.trigger = begin - self.trigger


        self.minutes = (self.trigger.seconds // 60) % 60
        self.hours = (self.trigger.seconds // 3600) % 60
        self.seconds = self.trigger.seconds - self.hours*3600 - self.minutes*60

        for i in ["AUDIO", "DISPLAY", "EMAIL"]:
            if i[0] == self.action.upper():
                self.action = i


    def print(self):
        trigger = f":-P{self.trigger.days}DT{self.hours}H{self.minutes}M{self.seconds}S"

        alarm = f"\nBEGIN:VALARM\nDESCRIPTION:{self.description}\nTRIGGER{trigger}\nACTION:{self.action}\nEND:VALARM"
        return(alarm)

def format_time(time):
    if isinstance(time, dt.datetime):
        return time.strftime("%Y%m%dT%H%M%S")
    else:
        return time.strftime("%Y%m%d")

def merge(list):
    merged_str = ""
    if len(list) > 0:
        for element in list:
            merged_str += element.print()
    return merged_str

if __name__ == "__main__":

    testcal = Calendar("Testkalender")

    begin = dt.datetime(2016, 7, 27, 00, 00, 00, 0)
    end = dt.datetime(2016, 7, 28, 00, 00, 00, 0)
    testevent1 = Event(organizer="Norman", summary="Jahrestag", description="Unser Jahrestag", begin=begin, end=end, busy=False, all_day=True, country_code="DE")
    testevent1.add_alarm(trigger=dt.timedelta(days=1), description="erste Erinnerung")
    testevent1.add_alarm(trigger=dt.datetime(2016, 7, 20, 22, 00, 00, 0), description="zweite Erinnerung")
    testevent1.set_feq_rule(freq="y")
    testcal.add_event(testevent1)

    testcal.save_ics()

