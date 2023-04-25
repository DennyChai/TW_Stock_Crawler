import DataTransformer
import CrawlerService
import DriverService
import DataProvider
import GoogleSheet
import traceback
import sys
import os


def send_msg(level: str = "Error", msg: str = "") -> None:
    sys_log = traceback.format_exc()
    send_message = msg if sys_log == "NoneType: None\n" else f"{sys_log}{msg}"
    print(f"開發測試: {send_message}")


def main():
    try:
        VERSION = "V1"
        PROJECT_PATH = os.path.abspath(os.getcwd())
        provider = DataProvider.DataProvider(send_msg, PROJECT_PATH)
        google_sheet = GoogleSheet.GoogleSheet(send_msg)
        transformer = DataTransformer.DataTransformer(send_msg, provider)
        driver_service = DriverService.DriverService(send_msg, provider, transformer)
        service_inputs = {
            "version": VERSION,
            "send_msg": send_msg,
            "provider": provider,
            "transformer": transformer,
            "input_data": sys.argv[1:],
            "google_sheet": google_sheet,
            "driver_service": driver_service,
        }
        CrawlerService.CrawlerService(service_inputs).stock_service()
    except:
        print("Main Error: ", traceback.format_exc())


if __name__ == "__main__":
    main()
