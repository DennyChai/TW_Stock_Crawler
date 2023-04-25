from copy import deepcopy
import AppSettings
import requests
import time


class DriverService(object):
    def __init__(self, send_msg, provider, transformer) -> None:
        self.send_msg = send_msg
        self.provider = provider
        self.transformer = transformer
        self.setting = deepcopy(AppSettings.settings["玩股網"])

    def main(self, sheet, start_date, end_date):
        try:
            driver = self.provider.get_driver()
            main_url = self.setting["main_url"]
            driver.get(main_url)
            time.sleep(5)
            headers = self.get_headers(driver)
            self.wantgoo_daily_stick(headers, start_date, sheet)
            self.wantgoo_pc_ratio(headers, start_date, sheet)
            self.wantgoo_big_8_trend(headers, start_date, sheet)

        except:
            self.send_msg()

    def wantgoo_big_8_trend(self, headers, end_date, sheet):
        try:
            setting = self.setting["八大行庫買賣動向"]
            pc_ratio_datas = self.provider.requests_data(setting["url"], format="json()", headers=headers)
            self.transformer.wantgoo_big_8_trend(sheet, pc_ratio_datas, setting, end_date)
        except:
            self.send_msg()

    def wantgoo_pc_ratio(self, headers, end_date, sheet):
        try:
            setting = self.setting["買賣權比"]
            pc_ratio_datas = self.provider.requests_data(setting["url"], format="json()", headers=headers)
            self.transformer.wantgoo_pc_ratio(sheet, pc_ratio_datas, setting, end_date)
        except:
            self.send_msg()

    def wantgoo_daily_stick(self, headers, start_date, sheet):
        # 日趨勢
        try:
            setting = self.setting["日趨勢"]
            yesterday_timestamp = self.transformer.get_timestamp(start_date)
            daily_stick_url = setting["url"].format(yesterday_timestamp)
            stick_datas = self.provider.requests_data(daily_stick_url, format="json()", headers=headers)
            self.transformer.wantgoo_daily_stick(sheet, stick_datas, setting)
        except:
            self.send_msg()

    def get_headers(self, driver):
        try:
            for request in driver.requests:
                if "/daily-candlesticks?" in request.url:
                    headers = dict(request.headers)
            return headers
        except:
            self.send_msg()
