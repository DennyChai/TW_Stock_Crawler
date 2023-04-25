from googleapiclient.discovery import Resource
from datetime import datetime, timedelta, time
from typing import Any
import AppSettings
import time as t


class DataTransformer(object):
    def __init__(self, send_msg, provider) -> None:
        self.send_msg = send_msg
        self.provider = provider
        self.sheet_id = AppSettings.settings["sheet_id"]
        self.setting = AppSettings.settings

    def get_yesterday_timestamp(self, days: int = 1):
        yesterday = datetime.now() - timedelta(days=days)
        time_stamp = round(datetime.combine(yesterday, time.min).timestamp()) * 1000
        return time_stamp

    def get_timestamp(self, date: str):
        _date = datetime.strptime(date, "%Y/%m/%d")
        previoud_date = _date - timedelta(days=1)
        time_stamp = round(datetime.combine(previoud_date, time.min).timestamp()) * 1000
        return time_stamp

    def get_date_from_timestamp(self, date):
        _date = datetime.fromtimestamp(date // 1000)
        return datetime.strftime(_date, "%Y/%m/%d")

    def string_num_to_int(self, value):
        try:
            return int(value.replace(",", ""))
        except:
            self.send_msg()

    def get_cell_data(self, sheet: Resource, range: str, default: Any = 0):
        value = default
        try:
            cell_data = sheet.values().get(spreadsheetId=self.sheet_id, range=range).execute()
            value_list = cell_data.get("values")
            if value_list is not None:
                value = value_list[0][0]
        except:
            self.send_msg()
        return value

    def future_transaction_row_data(self, sheet, datas, cache, transaction_sum, hold_sum):
        try:
            sheet_name = datas["work_sheet"]
            previous_hold = self.get_cell_data(sheet, datas["previous_hold"])
            previous_hold = self.string_num_to_int(previous_hold)
            previous_transaction = self.get_cell_data(sheet, datas["previous_transaction"])
            previous_transaction = self.string_num_to_int(previous_transaction)
            row_data = [
                cache["date"],
                cache["day"],
                cache["close"],
                cache["price_diff"],
                transaction_sum,
                transaction_sum - previous_transaction,
                hold_sum,
                hold_sum - previous_hold,
            ]
            self.insert_new_row(sheet, sheet_name, datas, row_data)
        except:
            self.send_msg()

    def copy_only(self, sheet: Resource, num: int = 1):
        try:
            for sheet_name, setting in self.setting["copy_only"].items():
                requests = []
                sheet_metadata = sheet.get(spreadsheetId=self.sheet_id).execute()
                sheets = sheet_metadata.get("sheets", "")
                sheet_names = [s for s in sheets if s["properties"]["title"] == sheet_name][0]
                sheet_id = sheet_names["properties"]["sheetId"]
                row = setting["start_row"]  # 要新增列的位置
                requests.append(
                    {
                        "insertDimension": {
                            "range": {
                                "sheetId": sheet_id,
                                "dimension": "ROWS",
                                "startIndex": row,
                                "endIndex": row + num,
                            },
                            "inheritFromBefore": False,
                        }
                    }
                )
                if "copy_cell_data" in setting:
                    requests.append(
                        {
                            "copyPaste": {
                                "source": {
                                    "sheetId": sheet_id,
                                    "startRowIndex": row + 1,
                                    "endRowIndex": row + 2,
                                    "startColumnIndex": setting["copy_cell_data"]["start_column"],
                                    "endColumnIndex": setting["copy_cell_data"]["end_column"],
                                },
                                "destination": {
                                    "sheetId": sheet_id,
                                    "startRowIndex": row,
                                    "endRowIndex": row + 1,
                                    "startColumnIndex": setting["copy_cell_data"]["start_column"],
                                    "endColumnIndex": setting["copy_cell_data"]["end_column"],
                                },
                                "pasteType": "PASTE_NORMAL",
                            }
                        }
                    )
                batch_update_request = {"requests": requests}
                sheet.batchUpdate(spreadsheetId=self.sheet_id, body=batch_update_request).execute()
        except:
            self.send_msg()

    def insert_new_row(self, sheet: Resource, sheet_name: str, setting: dict, row_data: list, num: int = 1) -> None:
        """插入新的列

        Args:
            sheet (googleapiclient.discovery.Resource): 試算表
            sheet_name (str): 表格的名稱
            setting (dict): AppSettings裡面的配置
            row_data (list): 要新增的資料
            num (int, optional): 要新增幾個列. Defaults to 1.
        """
        try:
            sheet_metadata = sheet.get(spreadsheetId=self.sheet_id).execute()
            sheets = sheet_metadata.get("sheets", "")
            sheet_names = [s for s in sheets if s["properties"]["title"] == sheet_name][0]
            sheet_id = sheet_names["properties"]["sheetId"]
            row = setting["start_row"]  # 要新增列的位置
            request = {
                "insertDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": row,
                        "endIndex": row + num,
                    },
                    "inheritFromBefore": False,
                }
            }
            request2 = {
                "pasteData": {
                    "data": ",".join(map(str, row_data)),
                    "type": "PASTE_NORMAL",
                    "delimiter": ",",
                    "coordinate": {"sheetId": sheet_id, "rowIndex": row, "columnIndex": 0},
                }
            }
            requests = [request, request2]
            if "copy_cell_data" in setting:  # 要拷貝公式&格式
                requests.append(
                    {
                        "copyPaste": {
                            "source": {
                                "sheetId": sheet_id,
                                "startRowIndex": row + 1,
                                "endRowIndex": row + 2,
                                "startColumnIndex": setting["copy_cell_data"]["start_column"],
                                "endColumnIndex": setting["copy_cell_data"]["end_column"],
                            },
                            "destination": {
                                "sheetId": sheet_id,
                                "startRowIndex": row,
                                "endRowIndex": row + 1,
                                "startColumnIndex": setting["copy_cell_data"]["start_column"],
                                "endColumnIndex": setting["copy_cell_data"]["end_column"],
                            },
                            "pasteType": "PASTE_NORMAL",
                        }
                    }
                )
            if "copy_cell_formula" in setting:  # 只拷貝公式
                requests.append(
                    {
                        "copyPaste": {
                            "source": {
                                "sheetId": sheet_id,
                                "startRowIndex": row + 1,
                                "endRowIndex": row + 2,
                                "startColumnIndex": setting["copy_cell_formula"]["start_column"],
                                "endColumnIndex": setting["copy_cell_formula"]["end_column"],
                            },
                            "destination": {
                                "sheetId": sheet_id,
                                "startRowIndex": row,
                                "endRowIndex": row + 1,
                                "startColumnIndex": setting["copy_cell_formula"]["start_column"],
                                "endColumnIndex": setting["copy_cell_formula"]["end_column"],
                            },
                            "pasteType": "PASTE_FORMULA",
                        }
                    }
                )
            batch_update_request = {"requests": requests}
            sheet.batchUpdate(spreadsheetId=self.sheet_id, body=batch_update_request).execute()
        except:
            self.send_msg(msg=f"Sheet:{sheet_name}, Row:{row}, Num:{num}")

    def future_big3(self, sheet, start_date, end_date):
        try:
            for title, datas in self.setting["三大法人期貨"].items():
                post_data = {"queryStartDate": start_date, "queryEndDate": end_date, "commodityId": datas["code"]}
                data_content = self.provider.requests_data(datas["url"], "post", post_data)
                data_content = data_content.split("\r\n")
                for data in data_content[1:]:
                    if not data:
                        continue
                    data_list = data.split(",")
                    if data_list[2] not in datas["work_sheet"]:
                        continue
                    sheet_name = datas["work_sheet"][data_list[2]]
                    self.insert_new_row(sheet, sheet_name, datas, data_list)
                t.sleep(3)
        except:
            self.send_msg()

    def option_big3(self, sheet, start_date, end_date):
        try:
            for title, datas in self.setting["選擇權"].items():
                post_data = {"queryStartDate": start_date, "queryEndDate": end_date, "commodityId": datas["code"]}
                data_content = self.provider.requests_data(datas["url"], "post", post_data)
                data_content = data_content.split("\r\n")
                for data in data_content[1:]:
                    if not data:
                        continue
                    data_list = data.split(",")
                    key = f"{data_list[3]}_{data_list[2]}"
                    if key not in datas["work_sheet"]:
                        continue
                    sheet_name = datas["work_sheet"][key]
                    self.insert_new_row(sheet, sheet_name, datas, data_list)
                t.sleep(3)
        except:
            self.send_msg()

    def future_transaction_details(self, sheet, start_date, end_date):
        # 期貨每日交易行情
        try:
            for title, datas in self.setting["期貨每日交易行情"].items():
                post_data = {
                    "queryStartDate": start_date,
                    "queryEndDate": end_date,
                    "commodity_id": datas["code"],
                    "down_type": datas["down_type"],
                }
                transaction_sum = 0  # 成交量
                hold_sum = 0  # 未沖銷量
                date_cache = ""  # 日期的快取，用來判斷是否換日期
                cache = {}  # 存放每個日期的總和資料

                data_content = self.provider.requests_data(datas["url"], "post", post_data)
                data_content = data_content.split("\r\n")
                for data in data_content[1:]:
                    if not data:
                        continue
                    data_list = data.split(",")
                    if "/" in data_list[2]:  # 價差行情表 不抓
                        continue
                    _date = data_list[0]

                    if _date != date_cache:
                        if cache:
                            self.future_transaction_row_data(sheet, datas, cache, transaction_sum, hold_sum)
                        transaction_sum = 0
                        hold_sum = 0
                        cache.clear()
                        date_cache = _date
                        _day = datetime.strptime(_date, "%Y/%m/%d").weekday()
                        cache.update(
                            {
                                "date": _date,
                                "day": self.setting["day_mapping"][_day],
                                "close": data_list[6],
                                "price_diff": data_list[7],
                            }
                        )
                        transaction_sum, hold_sum = self.future_transaction_details_calculation(
                            transaction_sum, hold_sum, data_list
                        )
                    else:
                        transaction_sum, hold_sum = self.future_transaction_details_calculation(
                            transaction_sum, hold_sum, data_list
                        )
                self.future_transaction_row_data(sheet, datas, cache, transaction_sum, hold_sum)
                t.sleep(3)
        except:
            self.send_msg()

    def future_transaction_details_calculation(self, transaction_sum, hold_sum, data_list):
        # 期貨每日交易行情 - 計算總和
        try:
            transaction_sum += int(data_list[9])  # 計算成交量
            if data_list[11] != "-":
                hold_sum += int(data_list[11])  # 計算未沖銷量
        except:
            self.send_msg()
        return transaction_sum, hold_sum

    def option_large_trader(self, sheet, start_date, end_date):
        # 選擇權大額交易人未沖銷部位
        try:
            for title, datas in self.setting["選擇權大額交易人未沖銷部位"].items():
                cache = []
                post_data = {"queryStartDate": start_date, "queryEndDate": end_date}
                data_content = self.provider.requests_data(datas["url"], "post", post_data)
                data_content = data_content.split("\r\n")
                for data in data_content[1:]:
                    if not data or "備註" in data:  # 後面幾行會有備註
                        continue
                    row_data = data.split(",")
                    key = f"{row_data[2]}{row_data[3]}"
                    if key not in datas["work_sheet"]:
                        continue
                    sheet_name = datas["work_sheet"][key]
                    if row_data[5] == "0":
                        cache.extend([row_data[0], row_data[2], row_data[3]])
                        if "999999" in row_data[4]:
                            cache.append("所有契約")
                        elif "666666" in row_data[4]:
                            cache.append("週契約")
                        else:
                            cache.append(row_data[4])
                        cache.extend(row_data[6:8])  # 五大 - 買方/賣方
                        cache.append(int(row_data[6]) - int(row_data[7]))
                        cache.extend(row_data[8:10])  # 十大 - 買方/賣方
                        cache.append(int(row_data[8]) - int(row_data[9]))
                    elif row_data[5] == "1":
                        cache.extend(row_data[6:8])  # 五特 - 買方/賣方
                        cache.append(int(row_data[6]) - int(row_data[7]))
                        cache.extend(row_data[8:10])  # 十特 - 買方/賣方
                        cache.append(int(row_data[8]) - int(row_data[9]))
                        self.insert_new_row(sheet, sheet_name, datas, cache)
                        cache.clear()
                t.sleep(3)
        except:
            self.send_msg()

    def future_large_trader(self, sheet, start_date, end_date):
        # 期貨大額交易人未沖銷部位
        try:
            for title, datas in self.setting["期貨大額交易人未沖銷部位"].items():
                cache = []
                post_data = {"queryStartDate": start_date, "queryEndDate": end_date}
                data_content = self.provider.requests_data(datas["url"], "post", post_data)
                data_content = data_content.split("\r\n")
                for sheet_name, key in datas["work_sheet"].items():
                    for data in data_content:
                        if key not in data:
                            continue
                        row_data = data.split(",")
                        if "999999" in row_data[3]:
                            if row_data[4] == "0":
                                cache.extend([row_data[0], "台指", "所有契約"])
                                up_yesterday = self.get_cell_data(sheet, f"{sheet_name}!D3")
                                up_yesterday = self.string_num_to_int(up_yesterday)
                                down_yesterday = self.get_cell_data(sheet, f"{sheet_name}!E3")
                                down_yesterday = self.string_num_to_int(down_yesterday)
                                dif1 = (down_yesterday - int(row_data[6])) - (up_yesterday - int(row_data[5]))
                                dif2 = int(row_data[5]) - int(row_data[6])
                                up_yesterday2 = self.get_cell_data(sheet, f"{sheet_name}!H3")
                                up_yesterday2 = self.string_num_to_int(up_yesterday2)
                                down_yesterday2 = self.get_cell_data(sheet, f"{sheet_name}!I3")
                                down_yesterday2 = self.string_num_to_int(down_yesterday2)
                                dif3 = (down_yesterday2 - int(row_data[8])) - (up_yesterday2 - int(row_data[7]))
                                dif4 = int(row_data[7]) - int(row_data[8])
                                cache.extend(
                                    [row_data[5], row_data[6], dif1, dif2, row_data[7], row_data[8], dif3, dif4]
                                )
                            elif row_data[4] == "1":
                                up_yesterday = self.get_cell_data(sheet, f"{sheet_name}!L3")
                                up_yesterday = self.string_num_to_int(up_yesterday)
                                down_yesterday = self.get_cell_data(sheet, f"{sheet_name}!M3")
                                down_yesterday = self.string_num_to_int(down_yesterday)
                                dif1 = (down_yesterday - int(row_data[6])) - (up_yesterday - int(row_data[5]))
                                dif2 = int(row_data[5]) - int(row_data[6])
                                up_yesterday2 = self.get_cell_data(sheet, f"{sheet_name}!P3")
                                up_yesterday2 = self.string_num_to_int(up_yesterday2)
                                down_yesterday2 = self.get_cell_data(sheet, f"{sheet_name}!Q3")
                                down_yesterday2 = self.string_num_to_int(down_yesterday2)
                                dif3 = (down_yesterday2 - int(row_data[8])) - (up_yesterday2 - int(row_data[7]))
                                dif4 = int(row_data[7]) - int(row_data[8])
                                cache.extend(
                                    [row_data[5], row_data[6], dif1, dif2, row_data[7], row_data[8], dif3, dif4]
                                )
                                self.insert_new_row(sheet, sheet_name, datas, cache)
                                cache.clear()
                t.sleep(3)
        except:
            self.send_msg()

    def wantgoo_daily_stick(self, sheet, datas, setting):
        # 日趨勢
        try:
            for daily_stick_data in reversed(datas):
                _date = self.get_date_from_timestamp(daily_stick_data["tradeDate"])
                open_price = round(daily_stick_data["open"])
                high_price = round(daily_stick_data["high"])
                close_price = round(daily_stick_data["close"])
                low_price = round(daily_stick_data["low"])
                volume = daily_stick_data["volume"]
                row_data = [_date, open_price, high_price, low_price, close_price, 0, volume]
                self.insert_new_row(sheet, setting["work_sheet"], setting, row_data)
        except:
            self.send_msg()

    def wantgoo_pc_ratio(self, sheet, datas, setting, end_date):
        # 買賣權比 P/C Ratio
        try:
            pc_ratio_data = datas[0]
            _date = self.get_date_from_timestamp(pc_ratio_data["date"])
            if _date == end_date:  # 確認是否為指定日期的資料
                call_vol = pc_ratio_data["callVolume"]  # 買權成交量
                put_vol = pc_ratio_data["putVolume"]  # 賣權成交量
                put_call_vol_ratio = pc_ratio_data["putCallVolumeRatio"]  # 成交量多空比%
                call_open_interest = pc_ratio_data["callOpenInterest"]  # 買權未平倉量
                put_open_interest = pc_ratio_data["putOpenInterest"]  # 賣權未平倉量
                put_call_interest_ratio = pc_ratio_data["putCallOpenInterestRatio"]  # 未平倉多空比%
                taiex_close_price = pc_ratio_data["taiexClose"]  # 加權指數收盤價
                tx_close_price = pc_ratio_data["txClose"]  # 台指收盤價
                row_data = [
                    _date,
                    put_vol,
                    call_vol,
                    put_call_vol_ratio,
                    put_open_interest,
                    call_open_interest,
                    put_call_interest_ratio,
                    taiex_close_price,
                    tx_close_price,
                ]
                self.insert_new_row(sheet, setting["work_sheet"], setting, row_data)
        except:
            self.send_msg()

    def wantgoo_big_8_trend(self, sheet, datas, setting, end_date):
        # 八大行庫買賣動向
        try:
            daily_data = datas[0]
            cache = {}
            _date = daily_data["date"][:10]
            if _date.replace("-", "/") == end_date:  # 確認是否為指定日期的資料
                row_data = [_date]
                for bank, bank_data in daily_data.items():
                    if bank == "date":
                        continue
                    amount = bank_data["amount"]
                    volume = bank_data["count"]
                    bank_name = self.setting["玩股網"]["bank_mapping"][bank]
                    cache[bank_name] = [volume, amount]
                for _, values in sorted(cache.items(), key=lambda x: x):
                    row_data.extend(values)
                self.insert_new_row(sheet, setting["work_sheet"], setting, row_data)
        except:
            self.send_msg()
