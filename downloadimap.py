import argparse
import email
import getpass
import imaplib
import os
import re
import socket
import traceback

import imapclient
import socks

parser = argparse.ArgumentParser("Download IMAP")
parser.add_argument("-s", "--server", type=str, required=True, help="Server Name")
parser.add_argument("-p", "--port", type=int, default=993, help="Server Port")
parser.add_argument("-e", "--email", type=str, required=True, help="Email Address")
parser.add_argument("-P", "--password", type=str, default=None, help="Password")
parser.add_argument("-o", "--output", type=str, default="download", help="Output Folder")
parser.add_argument("-x", "--proxy", type=str, default=None, help="HTTP Proxy Address")
parser.add_argument("-y", "--proxy_port", type=int, default=1080, help="HTTP Proxy Port")

args = parser.parse_args()

if args.password is None:
    args.password = getpass.getpass()

if args.proxy is not None:
    socks.setdefaultproxy(socks.PROXY_TYPE_HTTP, args.proxy, args.proxy_port, True)
    socket.socket = socks.socksocket

M = imaplib.IMAP4_SSL(args.server, port=args.port)

M.login(args.email, args.password)

SAVE_FOLDER = args.output

INVAILD_CHARS = set('\\/:*?"<>|"\'')
for i in range(32):
    INVAILD_CHARS.add(chr(i))


def getcodec(txt):
    if txt == "unknown-8bit":
        return "gbk"
    if txt:
        return txt
    return "utf8"


def escape_text(text):
    text = "".join((" " if x in INVAILD_CHARS else x for x in text))
    text = re.sub("\\s+", " ", text).strip()
    return text


def prase_text(text):
    text = email.header.decode_header(text)
    text = "".join(((x[0].decode(getcodec(x[1]), errors="ignore") if isinstance(x[0], bytes) else x[0]) for x in text))
    return escape_text(text)


def get_folders(M):
    folders = M.list()[1]
    for name in folders:
        name_select = name.decode("utf-8").split(' "/" ')[1].join(("", ""))
        name_show = imapclient.imap_utf7.decode(name).split(' "/" ')[1]
        print(name_show)
        yield name_show, name_select


def get_mails(*folders):
    for name_show, name_select in folders:
        print(name_select)
        M.select(name_select)
        try:
            mailids = M.search(None, "ALL")[1][0]
        except M.error:
            traceback.print_exc()
            continue

        if mailids is None:
            continue
        mailids = mailids.split()
        for mailid in mailids:
            mail1 = M.fetch(mailid, "BODY[HEADER]")[1][0][1]
            mail2 = M.fetch(mailid, "BODY[TEXT]")[1][0][1]
            mail = mail1 + mail2
            yield name_show, mail


for f, mail in get_mails(*get_folders(M)):
    mail_obj = email.message_from_bytes(mail)
    f = escape_text(f)

    # mailto = mail_obj.get_all('To')[0]
    # mailto = prase_text(mailfrom)

    # mailfrom = mail_obj.get_all('From')[0]
    # mailfrom = prase_text(mailfrom)

    timerecv = mail_obj.get_all("Date")
    if timerecv:
        timerecv = timerecv[0]
        timerecv = email.utils.parsedate_to_datetime(timerecv)
        timerecv = timerecv.strftime("%Y-%m-%d--%H-%M-%S")
    else:
        timerecv = "Unknown"

    mailsubject = mail_obj.get_all("Subject")
    mailsubject = mailsubject[0] if mailsubject else ""
    mailsubject = prase_text(mailsubject)

    filename = " ".join((timerecv.strip(), mailsubject.strip())) + ".eml"

    print(filename)

    output_path = os.path.join(SAVE_FOLDER, f)
    os.makedirs(output_path, exist_ok=True)

    outputfile = open(os.path.join(output_path, filename), "wb")
    outputfile.write(mail)
    outputfile.close()
