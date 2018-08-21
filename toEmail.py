import smtplib as sm
import time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import xlrd

EMAIL = '***********'  # 邮箱账号
password = '***********'  # 邮箱客户端授权码
efile = r'***********'  # Excel文件路径
uname = '***********'  # 正文、签名中发送者姓名
phone = '***********'  # 签名中联系电话
fj1file = r'***********'  # 附件1路径
fj2file = r'***********'  # 附件2路径
Bin = 0  # 开始行数
End = 999  # 结束行数


def mail(sjr, familName):
    ret = True
    try:
        Text = '邮件正文'
        text = Text
        msg = MIMEMultipart('mixed')
        text_plain = MIMEText(text, 'plain', 'utf-8')
        msg.attach(text_plain)
        print('正文')
        # print(text)

        # 附件1
        att1 = MIMEBase('application', 'octet-stream')
        att1.set_payload(open(fj1file, 'rb').read())
        att1.add_header('Content-Disposition', 'attachment', filename=('gbk', '', "***********"))  # 邮件中显示附件1名称
        encoders.encode_base64(att1)
        msg.attach(att1)
        # 附件2
        att2 = MIMEBase('application', 'octet-stream')
        att2.set_payload(open(fj2file, 'rb').read())
        att2.add_header('Content-Disposition', 'attachment', filename=('gbk', '', "***********"))  # 邮件中显示附件2名称
        encoders.encode_base64(att2)
        msg.attach(att2)
        print('附件')

        smtp = sm.SMTP()
        msg['Subject'] = '***********'  # 邮件主题
        msg['From'] = EMAIL
        msg['to'] = sjr
        smtp.connect('smtphm.qiye.163.com', '25')  # smtp服务器地址
        smtp.login(EMAIL, password)
        form_addr = EMAIL
        smtp.sendmail(EMAIL, sjr, msg.as_string())
        smtp.quit()
        print(sjr + familName)
    except Exception:
        ret = False
    else:
        print('发送成功' + sjr)
    return ret


# 获取邮箱以及称呼
def opex(i):
    excelFile = xlrd.open_workbook(efile)
    sheetname = excelFile.sheet_names()[0]
    sheet = excelFile.sheet_by_name('Sheet1')
    rows = sheet.row_values(i)
    return rows


def send():
    c = 0
    qs = 0
    sb = 0
    try:
        for i in range(Bin, End):
            print('---------------------------')
            list = opex(i)
            name = list[2]
            Nname = name.strip()
            e = list[6]
            if Nname != '' and e.strip() != '':
                print(Nname)
                bool = mail(e, Nname)
                if bool:
                    c = c + 1
                    print('%s %s' % ('本次运行已成功发送：', c))
                    print('休息5分钟')
                    time.sleep(300)
                else:
                    sb = sb + 1
                    print('%s %s' % ('发送失败：', sb))
            else:
                qs = qs + 1
                print('%s %s' % ('信息缺失发送失败', qs))
    except '':
        print('出错了')
    else:
        print('发送完成')


if __name__ == '__main__':
    send()
