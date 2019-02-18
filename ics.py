import datetime
import argparse


def parse_event(f):
    events = []
    for line in f:
        if line.count("BEGIN:VEVENT"):
            event = {"category":""}
            for line in f:
                line = line.strip()
                if 0 == line.count("END:VEVENT"):
                    if 0 < line.count("DTEND"):
                        event["end"] = line
                    if 0 < line.count("DTSTART"):
                        event["start"] = line
                    if 0 < line.count("CATEGORIES:"):
                        event["category"] = line.split(':')[1]
                else:
                    break
            events.append(event)
    return events


def get_args():

    parser = argparse.ArgumentParser(description='Calculate calendar allocation.')
    parser.add_argument(dest='file', help='an integer for the accumulator')
    parser.add_argument('-c', dest='measured_category', default='blocker', help='What CATEGORIES field to calculate')
    parser.add_argument('-e', dest='exclude', help='What CATEGORIES field to calculate')
    return parser.parse_args()


def to_datetime(string):
    time = string.split(':')[1]
    time = time.split('T')
    if len(time) == 2:
        time = time[1]
        return datetime.timedelta(hours=int(time[0:2]), minutes=int(time[2:4]))


if __name__ == "__main__":
    args = get_args()

    duration = datetime.timedelta()
    all_vents_duration = datetime.timedelta()

    f = open(args.file, errors="ignore")
    for event in parse_event(f):
        if args.exclude not in event["category"]:
            end = to_datetime(event["end"])
            start = to_datetime(event["start"])

            if end and start:
                all_vents_duration += end - start 
                if args.measured_category in event["category"]:
                    duration += end - start

    print('Duration of category "' + args.measured_category + '" is ' + str(duration) + " hours")

    seconds = all_vents_duration.total_seconds()
    minutes = seconds%3600/60
    hours = seconds/3600
    print("Total duration is " + str(int(hours)) + ":" + str(int(minutes)) + ":00" + " hours")
