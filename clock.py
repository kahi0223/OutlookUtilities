from datetime import datetime
from datetime import timedelta
from datetime import timezone


class Clock:
    """class of handling time
    """

    def today(self,
              tz: int = 9) -> datetime:
        """ get today in the specific timezone
        hour, minuite, secound == 0
        :param tz: timmezone(default = 9)
        :return: currennt time
        """
        tzinfo = timezone(timedelta(hours=tz))
        now = datetime.now(tzinfo)
        today = now.replace(hour=0, minute=0, second=0, microsecond=0)
        return today

    def get_tz_info(self, tz: int) -> timezone:
        """ convert time differance to timezone
        :param tz: time difference 
        :return: converted timezone
        """
        tzinfo = timezone(timedelta(hours=tz))
        return tzinfo

    def to_str(self, date: datetime, sep: str = '/') -> str:
        """ convert datetime to str
        :param datetime: datetime to convert
        :param sep: string delimiter. default is '/'
        :return: str of the time with delimiter
        """
        return date.strftime(r'%Y' + sep + r'%m' + sep + r'%d')

    def is_latest(self, date: str, tz: int) -> tuple:
        """ check the date is the latest(==today)
        :param date: target date to check
        :param tz: timezone number (UTC time difference)
        :retrun: Tuple of (bool, Optional[List[datetime]])
        return True if it is today, and None
        return False if it is not today, 
        and list of dates between target day and today   
        """
        today = Clock().today(tz=tz)

        date_diff = today - date
        if date_diff.days == 0:
            return True, None
        else:
            between_dates = []
            for i in range(date_diff.days):
                between_dates.append(date + timedelta(days=i+1))
            return False, between_dates


class TimeZone:
    def get_timezone(self, place: str=None) -> int:
        """ get timezone from str(class defined in config)
        """
        if place == 'China':
            tz = 8
        elif place == 'USA':
            tz = -8
        else:
            tz = 9
        return tz
