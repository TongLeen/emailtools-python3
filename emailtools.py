#################################################################################
#                              email tools                                      #
#                                                                               #
#           a simple wrapper for standard lib -- 'email' & 'stmplib'            #
#                                                                               #
#                                                                               #
#   name:     emailtools                                                        #
#   author:   TongLeen                                                          #
#   date:     2025-3-20                                                         #
#   version:  1.1.0                                                             #
#                                                                               #
#   python version:   3.12.2                                                    #
#   dependency:                                                                 #
#     - pillow (optional)                                                       #
#     - pandas (optional)                                                       #
#                                                                               #
#################################################################################
#                               how to use                                      #
# 1. Instantiate a 'Email' as mail                                              #
#   - use `mail = Email("mail_subject")` , it's the normal usage;               #
#   - use `mail = Email.sequence("mail_subject", <content1>, <ctnt2>,...)`      #
#       it's a convenient way to write a email.                                 #
#       content could be a 'str' or 'PIL.Image.Image';                          #
#       if str match markdown header, it will be inserted as a header:          #
#           e.g. `## the header` -> `<h2>the header</h2>`                       #
#       if pillow is avaliable, 'Image' can add to <content>                    #
#       if pandas is avaliable, 'DataFrame' can add to <content>                #
#                                                                               #
# 2. Fill it with your content                                                  #
#   For adding content after initialication, use <Email>.add* methods:          #
#   - addHead       : add a header in email; level means the header-level;      #
#   - addParagraph  : add a text paragraph                                      #
#   - addText       : add a text as header or paragraph                         #
#       if text match MD header, it will be converted to header; else paragraph #
#   - addTable      : add a list or list[list] as table                         #
#                                                                               #
#   Following methods are avaliable when specific package exists:               #
#   - addImage      : add a Image in email (avaliable when 'PIL' or 'pillow')   #
#   - addDataFrame  : add a Dataframe as table (avaliable when 'pandas')        #
#                                                                               #
# 3. Instantiate a 'EmailServer' as emailServer                                 #
#   - use `server = EmailServer(                                                #
#       "smtp_server_addr", "local_nickname", "local_account", "local_passwd")` #
#       it's the normal usage.                                                  #
#       !! Caution DO NOT write your account and passwd in code                 #
#   - use `server = EmainServer.initFromJson("config_file_path.json")`          #
#       this avoids your account and passwd appearing in code.                  #
#       !! The cfg file must be encoded by "utf-8"                              #
#                                                                               #
# 4. put the mail to emailServer                                                #
#   - use `server.send(mail, {"revc_1_addr@xxx.com":"his_nickname", ...})`      #
#   if 'server.defaultReveivers' specificed (set property | init from json)     #
#   - use `server.send(mail)` to send mail to all default receivers.            #
#                                                                               #
# Closing server manually is not necessary.                                     #
# It will be closed when server is deleted.                                     #
#                                                                               #
#################################################################################
#                                                                               #
# following example function shows usage:                                       #
#                                                                               #
# def example():                                                                #
#     img = Image.open("the_email_path")                                        #
#     mail = Email.sequence(                                                    #
#         "The subject",                                                        #
#         "# header-1",                                                         #
#         "paragraph",                                                          #
#         "another paragraph",                                                  #
#         img,                                                                  #
#         "# header-2",                                                         #
#         "another paragraph",                                                  #
#     )                                                                         #
#     server = EmailServer.initFromJson("./emailserver.json")                   #
#     server.send(mail)                                                         #
#                                                                               #
#################################################################################


CONFIG_TEMPLATE = """
{
    "host": "smtp.xxx.com",
    "addr": "email_address@xxx.com",
    "name": "your_nick_name",
    "key": "password",
    "receivers": {
        "reveiver_1_address@xx.com": "his/her_nickname"
    }
}
"""

import base64
import json
import re


from io import BytesIO
from smtplib import SMTP_SSL
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.header import Header
from typing import Self, Any, Sequence

try:
    from PIL import Image

    _USING_PIL = True
except ImportError:
    _USING_PIL = False
    print("Package 'PIL' not found. Relative functions not valid.")


try:
    import pandas as pd

    _USING_PANDAS = True
except ImportError:
    _USING_PANDAS = False
    print("Package 'pandas' not found. Relative functions not valid.")


HEAD_RE = re.compile(r"^\W*#+ .*$")
HTML_CSS = """
table {
  border-collapse: collapse;
  border: 2px solid rgb(140 140 140);
  font-family: sans-serif;
  font-size: 0.8rem;
  letter-spacing: 1px;
}
caption {
  caption-side: top;
  padding: 10px;
  font-weight: bold;
}
thead,
tfoot {
  background-color: rgb(228 240 245);
}
th,
td {
  border: 1px solid rgb(160 160 160);
  padding: 8px 10px;
  text-align: center;
}
tbody > tr:nth-of-type(even) {
  background-color: rgb(237 238 242);
}
"""


class Email:

    def __init__(self, subject: str):
        self.root = MIMEMultipart()
        self.elementList: list[str] = []
        self.imageList: list[tuple[int, bytes]] = []
        self.imageCount = 0

        self.subject = subject
        return

    def addHead(self, text: str, level: int = 1) -> None:
        self.elementList.append(f"<h{level}>{text}</h{level}>")
        return

    def addParagraph(self, text: str) -> None:
        self.elementList.append(f"<p>{text}</p>")
        return

    def addText(self, text: str) -> None:
        if HEAD_RE.match(text):
            head = text.lstrip()
            level = head.find(" ") + 1
            head = head[level:].lstrip()
            self.addHead(head, level)
        else:
            self.addParagraph(text)

    def addTable(
        self,
        content: Sequence[Sequence],
        head: Sequence | None = None,
        caption: str | None = None,
    ) -> None:
        e = []
        e.append('<table border="1" cellspacing="0" cellpadding="0">')
        if caption is not None:
            e.append(f"<caption>{caption}</caption>")
        if head is not None:
            h = [f"<td>{i}</td>" for i in head]
            e.extend(["<thead><tr>", *h, "</tr></thead>"])

        e.append("<tbody>")
        for row in content:
            r = [f"<td>{i}</td>" for i in row]
            e.extend(["<tr>", *r, "</tr>"])
        e.append("</tbody>")
        e.append("</table>")

        self.elementList.extend(e)
        return

    def _setSender(self, name: str, addr: str) -> None:
        self.root["From"] = Header(f"{self._encodeNikename(name)} <{addr}>", "ascii")
        return

    def _setReceivers(self, receivers: dict[str, str]) -> None:
        self.root["To"] = Header(
            ";".join(
                [
                    f"{self._encodeNikename(name)} <{addr}>"
                    for addr, name in receivers.items()
                ]
            ),
            "ascii",
        )
        return

    @staticmethod
    def _encodeNikename(name: str) -> str:
        if name.isascii():
            return name
        else:
            b = base64.b64encode(name.encode("utf-8"))
            return f'"=?utf-8?B?{b.decode("ascii")}?="'

    def toBytes(self) -> bytes:
        self.root["Subject"] = Header(self.subject, "utf-8")

        html = f'''<html>
        <head><style>{HTML_CSS}</style></head>
        <body>
        {'\n'.join(self.elementList)}
        </body></html>'''

        body = MIMEText(html, "html", "utf-8")
        self.root.attach(body)

        for num, data in self.imageList:
            img = MIMEImage(data)

            img.add_header("Content-ID", f"<img{num}>")
            self.root.attach(img)

        return self.root.as_bytes()

    @classmethod
    def sequence(cls, subject: str, *items: Any) -> Self:
        mail = cls(subject)
        for i in items:
            if isinstance(i, str):
                mail.addText(i)
            elif isinstance(i, Sequence):
                if bool(i) and isinstance(i[0], Sequence):
                    mail.addTable(i)
                else:
                    mail.addTable((i,))
            elif _USING_PIL and isinstance(i, Image.Image):
                mail.addImage(i)
            elif _USING_PANDAS and isinstance(i, pd.DataFrame):
                mail.addDataFrame(i)
            else:
                raise TypeError(f"This type {type(i)} cannot append to email.")
        return mail

    if _USING_PIL:

        def addImage(self, img: Image.Image) -> None:
            self.elementList.append(f'<p><img src="cid:img{self.imageCount}"/></p>')
            buf = BytesIO()
            img.save(buf, "png")
            self.imageList.append((self.imageCount, buf.getvalue()))
            self.imageCount += 1
            return

    if _USING_PANDAS:

        def addDataFrame(self, df: pd.DataFrame, caption: str | None = None) -> None:
            header = list(df.columns)
            body = df.values.tolist()
            self.addTable(body, head=header, caption=caption)
            return


class EmailServer:
    def __init__(self, host: str, userName: str, userAddr: str, key: str) -> None:
        self.host = host
        self.userName = userName
        self.userAddr = userAddr

        self.server = SMTP_SSL(host)
        self.server.login(userAddr, key)
        self.defaultReceivers = None
        return

    def send(self, mail: Email, receivers: dict[str, str] | None = None) -> None:
        mail._setSender(self.userName, self.userAddr)
        if receivers:
            pass
        elif bool(self.defaultReceivers):
            receivers = self.defaultReceivers
        else:
            raise ValueError("No receivers specified.")

        mail._setReceivers(receivers)
        self.server.sendmail(self.userAddr, receivers.keys(), mail.toBytes())
        return

    def __del__(self) -> None:
        self.server.close()
        return

    @classmethod
    def initFromJson(cls, filePath: str) -> Self:
        with open(filePath, "r", encoding="utf-8") as f:
            cfg: dict = json.load(f)
        try:
            server = cls(cfg["host"], cfg["name"], cfg["addr"], cfg["key"])
            server.defaultReceivers = cfg.pop("receivers")
        except IndexError as e:
            print("Configure file format cracked. Use below template:")
            print(CONFIG_TEMPLATE)
        return server
