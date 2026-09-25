import os
import sys
import time
from datetime import datetime
import requests
import win32com.client
from dotenv import load_dotenv

# Загрузка переменных окружения из .env
load_dotenv()

LOG_FILE = os.getenv("LOG_FILE", "ups_log.txt")


def write_to_log(message):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now().strftime('%d.%m.%Y %H:%M:%S')} - {message}\n")


# Считывание параметров из .env
try:
    TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
    TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
        raise ValueError("TELEGRAM_TOKEN или TELEGRAM_CHAT_ID не заданы!")

    CHECK_INTERVAL = float(os.getenv("CHECK_INTERVAL", 5))
    DELAY_NOTIFY = float(os.getenv("DELAY_NOTIFY", 60))
    SHUTDOWN_THRESHOLD = float(os.getenv("SHUTDOWN_THRESHOLD", 20))
    SHUTDOWN_TIMEOUT = float(os.getenv("SHUTDOWN_TIMEOUT", 30))

    if (
        CHECK_INTERVAL <= 0
        or DELAY_NOTIFY < 0
        or SHUTDOWN_THRESHOLD < 0
        or SHUTDOWN_TIMEOUT < 0
    ):
        raise ValueError("Числовые параметры должны быть неотрицательными!")

except Exception as e:
    error_msg = f"Ошибка конфигурации: {str(e)}"
    write_to_log(error_msg)
    sys.exit(1)


def send_to_telegram(text):
    try:
        requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage",
            data={"chat_id": TELEGRAM_CHAT_ID, "text": text},
            timeout=5,
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

            plugged = status == 2
            runtime = runtime if runtime != 0xFFFFFFFE else 0

            return percent, plugged, runtime

        write_to_log("Ошибка: батарея не обнаружена!")
        sys.exit(1)

    except Exception as e:
        write_to_log(f"Ошибка WMI: {str(e)}")
        sys.exit(1)


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

        if (
            was_on_battery
            and power_lost_time is not None
            and not telegram_notified
            and not plugged
        ):
            elapsed = int((datetime.now() - power_lost_time).total_seconds())
            if elapsed >= DELAY_NOTIFY:
                send_to_telegram(
                    f"{power_lost_time.strftime('%d.%m.%Y %H:%M:%S')} - Отключение электричества!\n"
                    f"Заряд: {charge_at_loss}%, осталось: {remaining_time_at_loss} мин."
                )
                telegram_notified = True

        if plugged and was_on_battery and power_lost_time is not None:
            restore_time = datetime.now()
            duration = int((restore_time - power_lost_time).total_seconds())
            log_msg = (
                f"Электричество восстановлено. Заряд: {charge}%, осталось: {remaining_time} мин, "
                f"прошло: {duration // 60} мин {duration % 60} сек."
            )
            write_to_log(log_msg)
            if duration >= DELAY_NOTIFY:
                send_to_telegram(
                    f"{restore_time.strftime('%d.%m.%Y %H:%M:%S')} - Электричество восстановлено.\n"
                    f"Заряд: {charge}%, осталось: {remaining_time} мин, прошла: {duration // 60} мин {duration % 60} сек."
                )
            sys.exit(0)

        if (
            was_on_battery
            and power_lost_time is not None
            and charge <= SHUTDOWN_THRESHOLD
        ):
            event_time = datetime.now()
            duration = int((event_time - power_lost_time).total_seconds())
            log_msg = (
                f"Выключение ПК! Заряд: {charge}%, осталось: {remaining_time} мин, "
                f"прошло: {duration // 60} мин {duration % 60} сек."
            )
            write_to_log(log_msg)
            send_to_telegram(
                f"{event_time.strftime('%d.%m.%Y %H:%M:%S')} - Выключение ПК!\n"
                f"Заряд: {charge}%, осталось: {remaining_time} мин, прошло: {duration // 60} мин {duration % 60} сек."
            )
            os.system(f"shutdown /s /t {int(SHUTDOWN_TIMEOUT)}")
            sys.exit(0)

        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(0)