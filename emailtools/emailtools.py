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
import sys

from collections.abc import Sequence
from io import BytesIO
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.header import Header
from smtplib import SMTP_SSL
from typing import Self, Any


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
        content: Sequence[Sequence[Any]],
        head: Sequence[Any] | None = None,
        caption: str | None = None,
    ) -> None:
        e: list[Any] = []
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
        self.root["From"] = Header(f"{self._encodeNikename(name)} <{addr}>", "ascii")  # type: ignore
        return

    def _setReceivers(self, receivers: dict[str, str]) -> None:
        self.root["To"] = Header(  # type: ignore
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
        self.root["Subject"] = Header(self.subject, "utf-8")  # type: ignore

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
        globalModules = sys.modules
        PIL = globalModules.get("PIL", None)
        pandas = globalModules.get("pandas")

        mail = cls(subject)
        for i in items:  # type: ignore
            i: "str | Sequence[Any] | Image | DataFrame"  # type: ignore
            if isinstance(i, str):
                mail.addText(i)
                continue
            if isinstance(i, Sequence):
                if any(isinstance(j, str) for j in i):  # type: ignore
                    mail.addTable((i,))  # type: ignore
                elif bool(i) and any(isinstance(j, Sequence) for j in i):  # type: ignore
                    mail.addTable(i)  # type: ignore
                else:
                    mail.addTable((i,))  # type: ignore
                continue

            if PIL is not None:
                if isinstance(i, PIL.Image.Image):
                    mail.addImage(i)  # type: ignore
                    continue
            if pandas is not None:
                if isinstance(i, pandas.DataFrame):
                    mail.addDataFrame(i)  # type: ignore
                    continue

            raise TypeError(f"This type {type(i)} cannot append to email.")  # type: ignore

        return mail

    def addImage(self, img: "Image") -> None:  # type: ignore
        self.elementList.append(f'<p><img src="cid:img{self.imageCount}"/></p>')
        buf = BytesIO()
        img.save(buf, "png")  # type: ignore
        self.imageList.append((self.imageCount, buf.getvalue()))
        self.imageCount += 1
        return

    def addDataFrame(
        self,
        df: "DataFrame",  # type: ignore
        caption: str | None = None,
    ) -> None:
        header = list(df.columns)  # type: ignore
        body = df.values.tolist()  # type: ignore
        self.addTable(body, head=header, caption=caption)  # type: ignore
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
        mail._setSender(self.userName, self.userAddr)  # type: ignore
        if receivers:
            pass
        elif bool(self.defaultReceivers):
            receivers = self.defaultReceivers
        else:
            raise ValueError("No receivers specified.")

        mail._setReceivers(receivers)  # type: ignore
        self.server.sendmail(self.userAddr, tuple(receivers.keys()), mail.toBytes())
        return

    def __del__(self) -> None:
        self.server.close()
        return

    @classmethod
    def initFromJson(cls, filePath: str) -> Self:
        with open(filePath, "r", encoding="utf-8") as f:
            cfg: dict[str, str] = json.load(f)
        try:
            server = cls(cfg["host"], cfg["name"], cfg["addr"], cfg["key"])
            server.defaultReceivers = cfg.pop("receivers")
            return server
        except IndexError as _:
            print("Configure file format cracked. Use below template:")
            print(CONFIG_TEMPLATE)
            sys.exit()
