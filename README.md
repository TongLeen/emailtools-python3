# email-tools

A wrapper for Python3 standard lib stmplib and email.

# 电子邮件工具

一个Python标准库 `smtp` 和 `email` 库的包装.

简化Python发送邮件的步骤，仅需几行代码即可生成格式工整的电子邮件并发送.

一个简单的示例：
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
以上代码即可实现邮件发送. 邮件内容可以有标题、文字和表格，如果有 `pillow` 和 `pandas` 库，也可以发送图片和 `DataFrame`.

### 配置文件

配置文件 `"./emailserver.json"` 包含发送邮件需要的信息，包括邮件服务器地址，发送邮件的账号密码，发件人名称，默认的邮件接收者（可选）. 配置文件内容如下：
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

## 使用方法

1. 新建 `Email` 实例. `Email` 类提供两种方式新建邮件:
```Python
mail_subject = "theSubjectOfMail"

# 通过实例化类来新建邮件
mail = Email(mail_subject)

# 通过类方法来新建邮件
mail = Email.sequence(mail_subject)
```
新建邮件时需要给定邮件的主题.

2. 向邮件添加内容. 添加内容可以通过 `mail.add*` 方法向邮件中添加内容
```Python
# 添加标题
mail.addHead("header1")             # 添加一级标题
mail.addHead("header2", level=1)    # 添加一级标题
mail.addHead("header3", level=2)    # 添加二级标题

# 添加正文
mail.addParagraph("This is a paragraph.")   # 添加正文段落

# 添加文本
# 使用这个方法可以实现添加标题和正文
# 标题使用 markdown 的标记
mail.addText("# header4")                   # 添加一级标题
mail.addText("This is another paragraph.")  # 添加正文段落

# 添加表格
t = [ [1,2,3], [4,5,6] ]    # 表格内容
th = ["A", "B", "C"]        # 表头
tc = "table_caption"        # 表格标题
mail.addTable(t, head=th, caption=tc) # 表头和标题是可选的

# 添加图片
from PIL import Image
img = Image.open("any.png")
mail.addImage(img)

# 添加DataFrame
import pandas as pd
df = pd.DataFrame(
    [ [1, 2, 3], [4, 5, 6] ], 
    columns=["A", "B", "C"]
)
mail.addDataFrame(df)
```
在新建邮件时，使用 `Email.sequence()` 方法时可以通过附加任意数量的参数向邮件添加内容:
```Python
mail = Email.sequence(
    "The subject",          # 邮件主题，必需参数
    "# header-1",           # 一级标题
    "paragraph",            # 段落
    "another paragraph",    
    img,                    # 图片
    "## header-2",          # 二级标题
    "another paragraph",
)
```
通过这种方式添加内容时，所有的内容自动按顺序添加到邮件当中. 其中的字符串会自动使用 `.addText` 进行处理，自动生成标题和正文.

3. 初始化邮件服务器. 初始化邮件服务器可以使用以下两种方式:

> **注意: 将账号信息放入代码中可能会泄漏；建议使用配置文件**
```Python
# 使用参数实例化
server = EmailServer(
    "smtp.xxx.com",             # 邮件服务器地址
    "local_nickname",           # 发信者昵称
    "local_address@xmail.com",  # 发信者账号
    "xxxxxxxxxxxxxxxxxxx",      # 发信者账号密码
)

# 使用配置文件
server = EmailServer.initFromJson("cfg.json")
```
配置文件格式见[这里](#配置文件).

4. 向邮件服务器发送邮件

```Python
# 将邮件交给服务器
server.send(mail)
```
邮件服务器在实例化后将保持登录状态. 当服务器实例被销毁时，自动退出登录. 

---
