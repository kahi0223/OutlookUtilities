import pandas as pd
import shutil
from datetime import datetime 
from datetime import timedelta
from clock import Clock

class CsvHander:
    """class to handle csv file
    """
    def __init__(self) -> None:
        pass

    def merge_csv(self,
                  main_csv_path: str,
                  new_csv_path: str) -> None:
        """ merge new csv file to main csv file
        overwrite to main csv file
        :param main_csv_path: path of main csv file(will overwrite)
        :param new_csv_path: path of csv file to merge
        """
        print("----Marging CSV")
        main_csv_pd = pd.read_csv(main_csv_path)

        new_csv_pd = pd.read_csv(new_csv_path)
        merged_pd = pd.concat([main_csv_pd, new_csv_pd], join = "outer", ignore_index = True)
        merged_pd.to_csv(main_csv_path, header = True, index = False,  mode ="w")
    
    def get_latest_day(self, csv_path: str, tz:int) -> datetime:
        """ get the last day of csv file (date of the last raw)
        """
        csv_df  = pd.read_csv(csv_path)
        tzinfo = Clock().get_tz_info(tz=tz)
        latest_day = datetime.strptime(csv_df.tail(1).Date.values[0], r"%Y/%m/%d").replace(tzinfo=tzinfo)
        return latest_day
    
    def copy_csv(self,
                 csv_path_from: str,
                 csv_path_to: str) -> None:
        """ copy csv (backup)
        """
        print("----Copying CSV")
        shutil.copy2(csv_path_from, csv_path_to)