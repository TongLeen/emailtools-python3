# emailtools-python3
                                                       
A simple wrapper for standard lib -- `email` & `stmplib`
  

Simplify the steps of sending emails in Python, generate and send neatly formatted emails with just a few lines of code.

A simple example：
``` Python
from PIL import Image
from emailtools import Email, EmailServer

img = Image.open("the_email_path")
mail = Email.sequence(
    "The subject",
    "# header-1",
    "paragraph", 
    "another paragraph",
    img,
    "# header-2",
    "another paragraph",
)
server = EmailServer.initFromJson("./emailserver.json")
server.send(mail)
```
The above code can be used to send emails. The content of the message can have headers, text, and tables. If `pillow` and `pandas` avaliable, image and  `DataFrame` can add to mail.

### config-file

Config file `"./emailserver.json"` (Of course you can use any name you like) contains mail-server-address, sender name, sender account and default receivers(optional). Format of config file is below：
```json
{
    "host": "smtp.xxx.com",
    "addr": "email_address@xxx.com",
    "name": "your_nick_name",
    "key": "password",
    "receivers": {
        "reveiver_1_address@xx.com": "his/her_nickname"
    }
}
```
---

## Dependency

- No essential dependency when running on Python3.12
    In this case, only text and table can add to mail
- `pillow` needed when sending image
- `pandas` needed when sending `DataFrame`


---

## How to use 

1. Create `Email` instance. `Email` provides two ways to create a mail:
```Python
mail_subject = "theSubjectOfMail"

# By instantiating `Email` class
mail = Email(mail_subject)

# By class method
mail = Email.sequence(mail_subject)
```
Subject of mail is needed when creating mail.

2. Add content to mail by `mail.add*` methods.
```Python
# add header
mail.addHead("header1")             # h1 (default)
mail.addHead("header2", level=1)    # h1
mail.addHead("header3", level=2)    # h2

# add paragraph
mail.addParagraph("This is a paragraph.")

# add any text
# start with one '#' or more would be seemed as header
mail.addText("# header4")                   # h1
mail.addText("This is another paragraph.")  # paragraph

# add table (2d)
t = [ [1,2,3], [4,5,6] ]    
th = ["A", "B", "C"]        
tc = "table_caption"
mail.addTable(t, head=th, caption=tc) # head ana caption is optional

# add Image (avaliable when `pillow`)
from PIL import Image
img = Image.open("any.png")
mail.addImage(img)

# add DataFrame (avaliable when `pandas`)
import pandas as pd
df = pd.DataFrame(
    [ [1, 2, 3], [4, 5, 6] ], 
    columns=["A", "B", "C"]
)
mail.addDataFrame(df)
```

More convenient, use classmethod `Email.sequence()` to create mail with any number of content:
```Python
mail = Email.sequence(
    "The subject",          # Subject (essential)
    "# header-1",           # h1
    "paragraph",            # p
    "another paragraph",    # p
    img,                    # image
    "## header-2",          # h2
    "another paragraph",    # p
)
```

3. Init `EmailServer` . Specific your account, passwd and mail-server-address:

> **Cation: DO NOT write your account in public code**
```Python
# use parameter
server = EmailServer(
    "smtp.xxx.com",             # address
    "local_nickname",           # local name
    "local_address@xmail.com",  # local account
    "xxxxxxxxxxxxxxxxxxx",      # local key
)

# use config file
server = EmailServer.initFromJson("cfg.json")
```
Format of config file see [this](#config-file).

4. Send the mail

```Python
# send the mail
server.send(mail)
```
`server` will retain logging in. Auto log out when instance deleted.
Manually log out is not needed.

---
