import time
import requests
import os
import sys
import configparser
from datetime import datetime
import win32com.client

# Чтение конфигурации из config.ini
config = configparser.ConfigParser()
try:
    with open("config.ini", encoding="utf-8") as f:
        config.read_file(f)
except FileNotFoundError:
    error_msg = "Ошибка: файл config.ini не найден!"
    temp_log = "ups_log.txt"
    with open(temp_log, 'a', encoding="utf-8") as f:
        f.write(f"{datetime.now().strftime('%d.%m.%Y %H:%M:%S')} - {error_msg}\n")
    sys.exit(1)
except UnicodeDecodeError as e:
    error_msg = f"Ошибка: не удалось декодировать config.ini: {str(e)}"
    temp_log = "ups_log.txt"
    with open(temp_log, 'a', encoding="utf-8") as f:
        f.write(f"{datetime.now().strftime('%d.%m.%Y %H:%M:%S')} - {error_msg}\n")
    sys.exit(1)

# Извлечение параметров из секции [Settings]
try:
    TELEGRAM_TOKEN = config["Settings"]["TELEGRAM_TOKEN"]
    TELEGRAM_CHAT_ID = config["Settings"]["TELEGRAM_CHAT_ID"]
    CHECK_INTERVAL = float(config["Settings"]["CHECK_INTERVAL"])
    DELAY_NOTIFY = float(config["Settings"]["DELAY_NOTIFY"])
    LOG_FILE = config["Settings"]["LOG_FILE"]
    SHUTDOWN_THRESHOLD = float(config["Settings"]["SHUTDOWN_THRESHOLD"])
    SHUTDOWN_TIMEOUT = float(config["Settings"]["SHUTDOWN_TIMEOUT"])
except KeyError as e:
    error_msg = f"Ошибка: параметр {str(e)} не найден в config.ini!"
    temp_log = config.get("Settings", {}).get("LOG_FILE", "ups_log.txt")
    with open(temp_log, 'a', encoding="utf-8") as f:
        f.write(f"{datetime.now().strftime('%d.%m.%Y %H:%M:%S')} - {error_msg}\n")
    sys.exit(1)
except ValueError as e:
    error_msg = f"Ошибка: неверное значение параметра в config.ini: {str(e)}"
    temp_log = config.get("Settings", {}).get("LOG_FILE", "ups_log.txt")
    with open(temp_log, 'a', encoding="utf-8") as f:
        f.write(f"{datetime.now().strftime('%d.%m.%Y %H:%M:%S')} - {error_msg}\n")
    sys.exit(1)

# Проверка параметров
if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
    error_msg = "Ошибка: TELEGRAM_TOKEN или TELEGRAM_CHAT_ID пустые!"
    temp_log = LOG_FILE if LOG_FILE else "ups_log.txt"
    with open(temp_log, 'a', encoding="utf-8") as f:
        f.write(f"{datetime.now().strftime('%d.%m.%Y %H:%M:%S')} - {error_msg}\n")
    sys.exit(1)
if CHECK_INTERVAL <= 0 or DELAY_NOTIFY < 0 or SHUTDOWN_THRESHOLD < 0 or SHUTDOWN_TIMEOUT < 0:
    error_msg = "Ошибка: CHECK_INTERVAL, DELAY_NOTIFY, SHUTDOWN_THRESHOLD и SHUTDOWN_TIMEOUT должны быть неотрицательными!"
    temp_log = LOG_FILE if LOG_FILE else "ups_log.txt"
    with open(temp_log, 'a', encoding="utf-8") as f:
        f.write(f"{datetime.now().strftime('%d.%m.%Y %H:%M:%S')} - {error_msg}\n")
    sys.exit(1)

# --- Функции ---

def write_to_log(message):
    with open(LOG_FILE, 'a', encoding="utf-8") as f:
        f.write(f"{datetime.now().strftime('%d.%m.%Y %H:%M:%S')} - {message}\n")

def send_to_telegram(text):
    try:
        requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage",
            data={"chat_id": TELEGRAM_CHAT_ID, "text": text},
            timeout=5
        )
    except requests.exceptions.RequestException as e:
        write_to_log(f"Ошибка Telegram: {str(e)}")

def get_battery_status():
    try:
        wmi = win32com.client.GetObject("winmgmts:")
        batteries = wmi.InstancesOf("Win32_Battery")
        
        for battery in batteries:
            percent = battery.EstimatedChargeRemaining
            status = battery.BatteryStatus
            runtime = battery.EstimatedRunTime
            
            if percent is None:
                write_to_log("Ошибка: процент заряда батареи не обнаружен!")
                sys.exit(1)
            
            plugged = (status == 2)
            runtime = runtime if runtime != 0xFFFFFFFE else 0
            
            return percent, plugged, runtime
        
        write_to_log("Ошибка: батарея не обнаружена!")
        sys.exit(1)
        
    except Exception as e:
        write_to_log(f"Ошибка WMI: {str(e)}")
        sys.exit(1)

# --- Основная логика ---

def main():
    power_lost_time = None
    was_on_battery = False
    telegram_notified = False
    remaining_time_at_loss = 0
    charge_at_loss = 0

    while True:
        charge, plugged, remaining_time = get_battery_status()

        if not plugged and not was_on_battery:
            was_on_battery = True
            power_lost_time = datetime.now()
            remaining_time_at_loss = remaining_time
            charge_at_loss = charge
            log_msg = f"Отключение электричества! Заряд: {charge_at_loss}%, осталось: {remaining_time_at_loss} мин."
            write_to_log(log_msg)

        if was_on_battery and not telegram_notified and not plugged:
            if (datetime.now() - power_lost_time).seconds >= DELAY_NOTIFY:
                send_to_telegram(
                    f"{power_lost_time.strftime('%d.%m.%Y %H:%M:%S')} - Отключение электричества!\n"
                    f"Заряд: {charge_at_loss}%, осталось: {remaining_time_at_loss} мин."
                )
                telegram_notified = True

        if plugged and was_on_battery:
            restore_time = datetime.now()
            duration = (restore_time - power_lost_time).seconds
            log_msg = f"Электричество восстановлено. Заряд: {charge}%, осталось: {remaining_time} мин, прошло: {duration // 60} мин {duration % 60} сек."
            write_to_log(log_msg)
            if duration >= DELAY_NOTIFY:
                send_to_telegram(
                    f"{restore_time.strftime('%d.%m.%Y %H:%M:%S')} - Электричество восстановлено.\n"
                    f"Заряд: {charge}%, осталось: {remaining_time} мин, прошло: {duration // 60} мин {duration % 60} сек."
                )
            sys.exit(0)

        if was_on_battery and charge <= SHUTDOWN_THRESHOLD:
            event_time = datetime.now()
            log_msg = f"Выключение ПК! Заряд: {charge}%, осталось: {remaining_time} мин, прошло: {(datetime.now() - power_lost_time).seconds // 60} мин {(datetime.now() - power_lost_time).seconds % 60} сек."
            write_to_log(log_msg)
            send_to_telegram(
                f"{event_time.strftime('%d.%m.%Y %H:%M:%S')} - Выключение ПК!\n"
                f"Заряд: {charge}%, осталось: {remaining_time} мин, прошло: {(datetime.now() - power_lost_time).seconds // 60} мин {(datetime.now() - power_lost_time).seconds % 60} сек."
            )
            os.system(f"shutdown /s /t {int(SHUTDOWN_TIMEOUT)}")
            sys.exit(0)

        time.sleep(CHECK_INTERVAL)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(0)
