import logging
import os
import sys
import time
from datetime import datetime
import requests
import win32com.client
from dotenv import load_dotenv

# Загрузка конфигурации
load_dotenv()

LOG_FILE = os.getenv("LOG_FILE", "ups_log.txt")

# Настройка логирования (в файл и в консоль)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%d.%m.%Y %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)

try:
    TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
    TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
        raise ValueError("TELEGRAM_TOKEN или TELEGRAM_CHAT_ID не заданы!")

    CHECK_INTERVAL = float(os.getenv("CHECK_INTERVAL", 5))
    DELAY_NOTIFY = float(os.getenv("DELAY_NOTIFY", 60))
    SHUTDOWN_THRESHOLD = float(os.getenv("SHUTDOWN_THRESHOLD", 20))
    SHUTDOWN_TIMEOUT = float(os.getenv("SHUTDOWN_TIMEOUT", 30))

except Exception as e:
    logging.critical(f"Ошибка конфигурации: {e}")
    sys.exit(1)


def send_to_telegram(text: str, retries: int = 3) -> bool:
    """Отправка сообщения в Telegram с повторными попытками при сбое сети."""
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {"chat_id": TELEGRAM_CHAT_ID, "text": text}

    for attempt in range(1, retries + 1):
        try:
            response = requests.post(url, data=payload, timeout=5)
            if response.status_code == 200:
                return True
            logging.warning(
                f"Telegram API вернул статус {response.status_code}: {response.text}"
            )
        except requests.exceptions.RequestException as e:
            logging.warning(
                f"Попытка {attempt}/{retries} отправки в Telegram не удалась: {e}"
            )
            time.sleep(2)
    return False


class BatteryMonitor:

    def __init__(self):
        try:
            self.wmi = win32com.client.GetObject("winmgmts:")
        except Exception as e:
            logging.critical(f"Не удалось инициализировать WMI: {e}")
            sys.exit(1)

    def get_status(self) -> tuple[int, bool, int]:
        try:
            batteries = self.wmi.InstancesOf("Win32_Battery")
            for battery in batteries:
                percent = battery.EstimatedChargeRemaining
                status = battery.BatteryStatus
                runtime = battery.EstimatedRunTime

                if percent is None:
                    logging.error("Процент заряда батареи не обнаружен!")
                    return 0, False, 0

                plugged = status == 2
                runtime = runtime if runtime != 0xFFFFFFFE else 0
                return percent, plugged, runtime

            logging.error("Батарея не обнаружена в системе!")
            return 0, False, 0
        except Exception as e:
            logging.error(f"Ошибка чтения WMI: {e}")
            return 0, False, 0


def main():
    monitor = BatteryMonitor()

    power_lost_time = None
    was_on_battery = False
    telegram_notified = False
    charge_at_loss = 0
    remaining_time_at_loss = 0

    logging.info("Мониторинг UPS успешно запущен.")

    while True:
        charge, plugged, remaining_time = monitor.get_status()

        # 1. Фиксация момента отключения питания
        if not plugged and not was_on_battery:
            was_on_battery = True
            power_lost_time = datetime.now()
            charge_at_loss = charge
            remaining_time_at_loss = remaining_time
            logging.warning(
                f"Отключение электричества! Заряд: {charge_at_loss}%, осталось: {remaining_time_at_loss} мин."
            )

        # 2. Отправка уведомления с задержкой (защита от скачков)
        if (
            was_on_battery
            and power_lost_time
            and not telegram_notified
            and not plugged
        ):
            elapsed = int((datetime.now() - power_lost_time).total_seconds())
            if elapsed >= DELAY_NOTIFY:
                msg = (
                    f"{power_lost_time.strftime('%d.%m.%Y %H:%M:%S')} — Отключение электричества!\n"
                    f"Заряд: {charge_at_loss}%, осталось: {remaining_time_at_loss} мин."
                )
                if send_to_telegram(msg):
                    telegram_notified = True

        # 3. Восстановление питания (сброс состояния вместо sys.exit)
        if plugged and was_on_battery and power_lost_time:
            restore_time = datetime.now()
            duration = int((restore_time - power_lost_time).total_seconds())
            log_msg = (
                f"Электричество восстановлено. Заряд: {charge}%, осталось: {remaining_time} мин, "
                f"прошло: {duration // 60} мин {duration % 60} сек."
            )
            logging.info(log_msg)

            if duration >= DELAY_NOTIFY:
                send_to_telegram(
                    f"{restore_time.strftime('%d.%m.%Y %H:%M:%S')} — Электричество восстановлено.\n"
                    f"Заряд: {charge}%, прошло: {duration // 60} мин {duration % 60} сек."
                )

            # Сброс флагов для продолжения мониторинга
            was_on_battery = False
            telegram_notified = False
            power_lost_time = None

        # 4. Критический разряд — выключение ПК
        if (
            was_on_battery
            and power_lost_time
            and charge > 0
            and charge <= SHUTDOWN_THRESHOLD
        ):
            event_time = datetime.now()
            duration = int((event_time - power_lost_time).total_seconds())
            msg = (
                f"{event_time.strftime('%d.%m.%Y %H:%M:%S')} — Критический уровень заряда ({charge}%)!\n"
                f"Инициировано выключение ПК."
            )
            logging.critical(msg)
            send_to_telegram(msg)

            # Выключение системы
            os.system(f"shutdown /s /t {int(SHUTDOWN_TIMEOUT)}")
            sys.exit(0)

        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logging.info("Мониторинг остановлен пользователем.")
        sys.exit(0)