from datetime import datetime
from typing import Tuple
import AppSettings


class CrawlerService(object):
    def __init__(self, service_inputs) -> None:
        self.version = service_inputs["version"]
        self.send_msg = service_inputs["send_msg"]
        self.provider = service_inputs["provider"]
        self.input_data = service_inputs["input_data"]
        self.transformer = service_inputs["transformer"]
        self.google_sheet = service_inputs["google_sheet"]
        self.driver_service = service_inputs["driver_service"]
        self.setting = AppSettings.settings
        self.date = ""

    def stock_service(self):
        start_date, end_date = self.get_date_list()
        print(f"日期區間 :{start_date} -  {end_date}")
        sheet = self.google_sheet.get_sheet()
        self.transformer.copy_only(sheet)  # 不須爬資料，只要拷貝公式即可
        self.transformer.future_large_trader(sheet, start_date, end_date)  # 期貨大額交易人未沖銷部位
        self.transformer.option_large_trader(sheet, start_date, end_date)  # 選擇權大額交易人未沖銷部位
        self.transformer.future_transaction_details(sheet, start_date, end_date)  # 期貨每日交易行情
        self.transformer.future_big3(sheet, start_date, end_date)  # 三大法人期貨
        self.transformer.option_big3(sheet, start_date, end_date)  # 三大法人選擇權
        self.provider.close_session()  # 這樣才不會共用到同個Session
        self.driver_service.main(sheet, start_date, end_date)

    def get_date_list(self) -> Tuple[str, str]:
        """依照使用者輸入的爬蟲類別取得日期

        Returns:
            Tuple[str, str]: 日期. Ex: '2022/12/12', '2022/12/13'
        """
        try:
            if self.input_data[0] == "0":
                date_now = datetime.strftime(datetime.now(), "%Y/%m/%d")
                if self.date != date_now:
                    self.date = date_now
                start_date = end_date = date_now
            elif self.input_data[0] == "1":
                start_date = end_date = self.input_data[1]
            elif self.input_data[0] == "2":
                start_date = self.input_data[1]
                end_date = self.input_data[2]
        except:
            self.send_msg()
        return start_date, end_date
