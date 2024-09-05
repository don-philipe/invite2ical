import configparser
import email
from icalendar import Calendar
import imaplib
import os.path


def check_mails(user, passwd, server, sender: str, n: int, mailbox: str) -> []:
    """
    Check last n mails from sender list with ics attachments. Gives back a list of attachments.
    :param user: email username
    :param passwd: email password
    :param server: email imap server
    :param sender: the email adresse which should be checked
    :param n: number of mails that should be checked
    :return: a list with found ics attachments from the named senders
    """
    attachments = []
    imap = imaplib.IMAP4_SSL(server)
    imap.login(user, passwd)

    status, messages = imap.select(mailbox)
    msgs = int(messages[0])
    for i in range(msgs, msgs - n, -1):
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                m = email.message_from_bytes(response[1])
                from_str, encoding = decode(m.get("From"))
                if from_str == sender and m.is_multipart():
                    for part in m.walk():
                        content_disposition = str(part.get("Content-Disposition"))
                        if "attachment" in content_disposition:
                            ics = part.get_payload(decode=True)
                            attachments.append(ics)

    imap.close()
    imap.logout()
    return attachments


def generate_calendar(attachments: [], replace_summary: str, calendar=None):
    """
    Generate a calendar object out of overhanded ics attachments. Takes only VEVENT components into account. When
    calendar is given, check for every event if it is already present in that calendar (via UID). Events will be
    updated when the LAST-MODIFIED timestamp is newer.
    :param attachments: a list of ics
    :param replace_summary: a string to replace the summary in all events
    :param calendar: a pre-existing calendar object to add events to, None by default
    :return: a calendar object with all events from attachments
    """
    if calendar is None:
        calendar = Calendar()

    uids = {}
    for event in calendar.walk(name="VEVENT"):
        uids[event['UID']] = event['LAST-MODIFIED']

    for attachment in attachments:
        cal = Calendar.from_ical(attachment)
        for event in cal.walk(name="VEVENT"):
            event['SUMMARY'] = replace_summary
            if event['UID'] not in uids.keys():
                # just add, because UID doesn't exist in calendar yet
                calendar.add_component(event)
            elif str(event['LAST-MODIFIED']) > str(uids[event['UID']]):
                # replace event with a newer version
                for idx, ev in enumerate(calendar.subcomponents):
                    if ev['UID'] == event['UID']:
                        calendar.subcomponents.pop(idx)
                        calendar.add_component(event)

    return calendar


def read_calendar(filename: str):
    """
    Check if calendar file exists. If it does, read it and return a calendar object. If it doesn't exist, create new
    calendar object.
    :param filename: the filename of the file to read
    :return: a calendar object, empty if file doesn't exists
    """
    if os.path.isfile(filename):
        with open(filename, "r") as f:
            content = f.readlines()
        return Calendar.from_ical("".join(content))
    else:
        return Calendar()


def write_calendar(filename: str, calendar):
    """
    Write calendar to file.
    :param filename: the name of the file to write
    :param calendar: the calendar object to write to file
    """
    with open(filename, "wb") as f:
        f.write(calendar.to_ical())


def decode(part):
    """
    Decode an email part like 'subject' or 'From'.
    :param part:
    :return: tuple with decoded string and charset
    """
    string, encoding = email.header.decode_header(part)[0]
    if isinstance(string, bytes):
        string = string.decode(encoding)
    return string, encoding


config = configparser.ConfigParser()
config.read("config.toml")
a = check_mails(config["invite2ical.email"]["username"],
                config["invite2ical.email"]["password"],
                config["invite2ical.email"]["server"],
                config["invite2ical.email"]["sender"],
                int(config["invite2ical.email"]["fetch_number"]),
                config["invite2ical.email"]["mailbox"])
filename = config["invite2ical"]["icsfile"]
rc = read_calendar(filename)
c = generate_calendar(a, config["invite2ical"]["summary"], rc)
write_calendar(filename, c)
