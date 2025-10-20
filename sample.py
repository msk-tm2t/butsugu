"""
    Python Liblary
"""
from pathlib import Path
import subprocess
import time
from datetime import datetime
import winreg
import xlwings as xw
from xlwings.constants import WindowState
from typing import Tuple
import os
import csv
import winsound
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.schedulers.blocking import BlockingScheduler
#from apscheduler.jobstores.mongodb import MongoDBJobStore
#from apscheduler.jobstores.sqlalchemy import SQLAlchemyJobStore
from apscheduler.jobstores.memory import MemoryJobStore
from apscheduler.executors.pool import ThreadPoolExecutor, ProcessPoolExecutor
import pythoncom
import psutil

""""
    Excelの実行ファイルのパスをWindowsのレジストリから取得する関数
"""
def get_path_to_xl() -> Path:
    path = r'SOFTWARE\MICROSOFT\Windows\CurrentVersion\App Paths\Excel.exe'
    key = winreg.OpenKeyEx(winreg.HKEY_LOCAL_MACHINE, path)
    """
    Excelのパスをレジストリから取得
    """
    data, _ = winreg.QueryValueEx(key, 'Path')
    return Path(rf'{data}\EXCEL.EXE')

""""
    Excelを立ち上げる関数
"""
def add_xl_app(add_book: bool=True) -> xw.App:
    """
        Excelを起動する
    """
    command = f'"{str(get_path_to_xl())}" /x /e'
    proc = subprocess.Popen(command)
    """
        Excelが起動が完了するまで待つ
    """
    while True:
        try:
            xl_app = xw.apps[proc.pid]
            break
        except:
            time.sleep(0.5)
    """
        Excelのブックを開く
    """
    if add_book:
        xl_app.books.add()
    return xl_app

def openExcel(xlfile: str=, shname: str) -> Tuple[xw.App, xw.Book, xw.Sheet] :
    while True :
        app = add_xl_app(False)
        flg, wb, wsrss = open_xl(app, xlfile, shname)
        if flg :
            break
    time.sleep(15)
    return app, wb, wsrss







if __name__ == '__main__':

    """
    スクリプト名を表示
    """
    print('Script: ', __file__, 'PID: ', psutil.Process().pid, 'Priority: ', psutil.Process().nice())

    # 初期設定
    xlfile = os.path.join(os.path.dirname(__file__), filename)
    shname = r'楽天RSS'

    """
    株価情報ファイルをオープンする
    """
    app, wb, wsrss = openExcel(xlfile, shname)

    """
        配列を追加
    """
    try:
        data = wb.sheets[shname].range('A1:EU201').value
    except Exception as e:
        print(e)
    else:
        csvfile = os.path.join('c:\\Data\\', datetime.now().strftime('%Y%m%d-%H%M%S') + '.csv')
        with open(csvfile, mode='w', encoding='UTF-8', newline='') as fcsv :
            writer = csv.writer(fcsv)
            writer.writerows(data)
    finally:
        """
            終了処理
             Excelファイルを保存して終了
        """
        time.sleep(1)
        wb.save()
        time.sleep(1)
        wb.close()
        time.sleep(1)
        app.quit()
