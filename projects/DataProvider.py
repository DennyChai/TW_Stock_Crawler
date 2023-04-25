from selenium.webdriver.support.ui import WebDriverWait as wait, Select
from selenium.webdriver.support import expected_conditions as EC
from requests.exceptions import ConnectTimeout, ConnectionError
from selenium.webdriver.chrome.service import Service
from selenium.webdriver import DesiredCapabilities
from datetime import datetime, time, timedelta
from selenium.webdriver.common.by import By
from seleniumwire import webdriver
from typing import Any
import AppSettings
import requests
import tempfile
import zipfile
import os
import re


class DataProvider(object):
    def __init__(self, send_msg, project_path) -> None:
        self.session = None
        self.headers = AppSettings.settings["headers"]
        self.project_path = project_path
        self.selenium_path = project_path + "\\WantgooDriver.exe"
        self.send_msg = send_msg

    def requests_url(self, url: str, method: str, post_data: Any, headers) -> requests.models.Response:
        """
        請求網址

        Args:
            url (str): 網址
            method (str): 方法
            post_data (Any): Post要帶的資料

        Returns:
            requests.models.Response: 請求的結果
        """
        try:
            session = self.get_session(headers)
            if method == "get":
                response = session.get(url, timeout=60)
            else:
                if post_data:
                    response = session.post(url, data=post_data, timeout=60)
                else:
                    response = session.post(url, timeout=60)
            return response
        except:
            self.send_msg(msg=f"Url:{url}, Method:{method}, Post Data:{post_data}")

    def requests_data(
        self, url: str, method: str = "get", post_data: Any = None, format: str = "text", headers=None
    ) -> Any:
        """
        將請求的資料轉換成format的格式並回傳

        Args:
            url (str): 請求的網址
            method (str, optional): 方法. Defaults to "get".
            format (str, optional): 要轉換的格式. Defaults to "json()".
            post_data (Any, optional): 方法是Post時要帶的資料. Defaults to [].

        Returns:
            Any: 轉換過的資料, 失敗時會回傳None
        """
        status_code = None
        try:
            response = self.requests_url(url, method, post_data, headers)
            status_code = response.status_code
            return eval(f"response.{format}")
        except:
            self.close_session()
            msg = f"url: {url}, method:{method}, format:{format}, post_data:{post_data}, status_code:{status_code}, response:{response.text},"
            self.send_msg(msg=msg)
            return None

    def get_session(self, headers) -> requests.Session:
        if self.session is None:
            self.session = requests.Session()
            if headers is None:
                self.session.headers.update(self.headers)  # 初始Headers
            else:
                self.session.headers.update(headers)  # 自訂需要的Headers

        return self.session

    def close_session(self) -> None:
        if self.session is not None:
            self.session.close()
            self.session = None

    def get_driver(self) -> webdriver:
        try:
            capabilities = DesiredCapabilities.CHROME
            capabilities["goog:loggingPrefs"] = {"performance": "ALL"}
            options = webdriver.ChromeOptions()
            options.add_argument("--incognito")  # 無痕設定
            options.add_argument("--disable-gpu")  # 規避google bug
            options.add_argument("ignore-certificate-errors")
            options.add_argument("--disable-blink-features=AutomationControlled")  # 是否driver控制瀏覽器
            options.add_experimental_option("excludeSwitches", ["enable-logging"])  # 關閉除錯LOG(selenium自帶的,並非程式碼錯誤)
            options.add_experimental_option("excludeSwitches", ["enable-automation"])  # 開發者模式
            webdriver_path = Service("WantgooDriver.exe")
            driver = self.check_driver_version(capabilities, options, webdriver_path)
            driver.set_window_size(1100, 1150)  # 視窗大小
            return driver
        except:
            self.send_msg()

    def check_driver_version(
        self, capabilities: dict, options: webdriver.ChromeOptions, webdriver_path: Service
    ) -> webdriver:
        """
        檢查Selenium Driver的版本

        Args:
            capabilities (dict): Selenium Driver的功能
            options (webdriver.ChromeOptions): Selenium Driver的屬性
            webdriver_path (Service): Selenium Driver的路徑

        Returns:
            webdriver: Selenium Driver
        """
        try:
            driver = webdriver.Chrome(desired_capabilities=capabilities, options=options, service=webdriver_path)
            return driver
        except Exception as e:
            self.send_msg()
            url = self.get_version_url(str(e))
            self.download(url)
            os._exit(0)

    def get_version_url(self, err_msg: str) -> str:
        """
        取得要下載的連結

        Args:
            err_msg (str): 錯誤訊息

        Returns:
            str: 下載連結
        """
        try:
            down_load_html = requests.get("https://chromedriver.chromium.org/downloads").text
            url_list = re.findall(r"https://chromedriver.storage.googleapis.com/index.html.*?/", down_load_html)
            if "executable needs to be in PATH" in err_msg:
                down_url = url_list[0].replace("index.html?path=", "") + "chromedriver_win32.zip"
            elif "version is" in err_msg:  # 無version字串就是非版本錯誤,無法自動下載,需另外查看
                version = "".join(re.findall(r"version is (.*?) with", err_msg))
                version = ".".join(version.split(".")[:3])
                for url in url_list:
                    if version in url:
                        down_url = url.replace("index.html?path=", "") + "chromedriver_win32.zip"
            print("FINAL URL", down_url)
            return down_url
        except:
            self.send_msg()

    def download(self, url: str) -> None:
        """
        下載Selenium Driver, 並解壓縮改名放到指定路徑

        Args:
            url (str): 下載連結
        """
        try:
            driver_exe = requests.get(url).content
            tmp_file = tempfile.TemporaryFile()
            tmp_file.write(driver_exe)
            zf = zipfile.ZipFile(tmp_file, mode="r")
            for name in zf.namelist():
                f = zf.extract(name, self.project_path)
                if "chromedriver.exe" not in f:
                    os.remove(f)
                    continue
                elif os.path.isfile(self.selenium_path):
                    os.remove(self.selenium_path)
                os.rename(f, self.selenium_path)
            zf.close()
        except:
            self.send_msg()
