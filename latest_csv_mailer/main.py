import sys
from emailer import Emailer
from config import MailConfig
from config import PathConfig
from csv_hander import CsvHander
from clock import TimeZone
from clock import Clock
import os

def main(subject:str, place:str=None) -> None:
    """ to update(merge) csv file if latest csv file is sent by outlook
    serch and save attached the latest csv file of email and merge main csv file
    """

    main_csv_path = os.path.join(PathConfig.DIR_PATH, 
                            PathConfig.CSV_PATH)
    backup_csv_path = os.path.join(PathConfig.DIR_PATH, 
                            PathConfig.BACKUP_CSV_PATH)
    tz = TimeZone().get_timezone(place=place)

    last_day = CsvHander().get_latest_day(csv_path=main_csv_path, tz=tz)
    is_latest, between_dates =Clock().is_latest(date=last_day, tz=tz)
    if not is_latest:
        new_csv_paths = []

        # search email with old dates
        for date in between_dates:
            attachments = Emailer().search_mail_and_save_attachments(
                subject = MailConfig.GET_MAIL_SUBJECT.format(
                    subject=subject, date=date),
                save_path = os.path.join(PathConfig.DIR_PATH, \
                    PathConfig.RAW_CSV_PATH.format(date=date)),
                folder_name = MailConfig.GET_MAIL_FOLDER,
                )
            if attachments:
                for attachment in attachments:
                    new_csv_paths.append(attachment)
                last_day = date

        # merge latest csv file to main csv
        if new_csv_paths:
            for new_csv_path in new_csv_paths:
                CsvHander().merge_csv(csv_path=main_csv_path,
                                      new_csv_path=new_csv_path)
        
            CsvHander().copy_csv(csv_path_from=main_csv_path, csv_path_to=backup_csv_path)


    Emailer().create_new_mail(subject = MailConfig.SEND_MAIL_SUBJECT.format(subject=subject, date=last_day),
                              body = MailConfig.BODY.format(path=main_csv_path),                           
                              to = MailConfig.get_to(),
                              cc = MailConfig.get_cc(),
                              attach_paths = [main_csv_path])



if __name__ == "__main__":
    subject = sys.argv[1]
    main(subject=subject)


    