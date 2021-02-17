from datetime import datetime

class MailConfig:

    GET_MAIL_FOLDER = 'SubFolder'
    GET_MAIL_SUBJECT = 'LatestData_{subject}_{date:%Y/%m/%d(%Z)}' # BaseSubject_{sub}_2021/01/04(UTC+08:00)

    SEND_MAIL_SUBJECT = 'LatestMerged_{subject}_{date:%Y/%m/%d(%Z)}'

    TO1 = ["aaa@XXX.com", "bbb@XXX.com"]
    TO2 = ["aaa@XXX.com", "bbb@XXX.com"]

    CC1 = ["AAA@XXX.com", "BBB@XXX.com"]
    CC2 = ["AAA@XXX.com", "BBB@XXX.com"]

    BODY = 'Dear XXX,<br><br> \
        XXXXX.<br> \
        SAVED folder <br> ({path})<br><br>'
    
    @classmethod
    def get_to(cls, to:str="All") -> list:
        if to == '1':
            return cls.TO1
        elif to == '2':
            return cls.TO2
        else:
            return cls.TO1 + cls.TO2

    @classmethod
    def get_cc(cls, cc:str="All") -> list:
        if cc == '1':
            return cls.CC1
        elif cc == '2':
            return cls.CC2
        else:
            return cls.CC1 + cls.CC2

class PathConfig:
    DIR_PATH = r'C:\Users\FOLDERNAME' 

    CSV_PATH = r'CSV\MAIN.csv'
    BACKUP_CSV_PATH = r'CSV\MAIN_BACKUP.csv'

    RAW_CSV_PATH = r'CSV\RAW\RAW_{date:%y-%m-%d}.csv'
