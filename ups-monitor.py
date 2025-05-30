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
    if not config.read("config.ini"):
        raise FileNotFoundError("Файл config.ini не найден!")
except FileNotFoundError:
    error_msg = "Ошибка: файл config.ini не найден!"
    temp_log = "ups_log.txt"  # Временный лог, если LOG_FILE не задан
    with open(temp_log, 'a') as f:
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
    with open(temp_log, 'a') as f:
        f.write(f"{datetime.now().strftime('%d.%m.%Y %H:%M:%S')} - {error_msg}\n")
    sys.exit(1)
except ValueError as e:
    error_msg = f"Ошибка: неверное значение параметра в config.ini: {str(e)}"
    temp_log = config.get("Settings", {}).get("LOG_FILE", "ups_log.txt")
    with open(temp_log, 'a') as f:
        f.write(f"{datetime.now().strftime('%d.%m.%Y %H:%M:%S')} - {error_msg}\n")
    sys.exit(1)

# Проверка положительных числовых параметров
if CHECK_INTERVAL <= 0 or DELAY_NOTIFY < 0 or SHUTDOWN_THRESHOLD < 0 or SHUTDOWN_TIMEOUT < 0:
    error_msg = "Ошибка: CHECK_INTERVAL, DELAY_NOTIFY, SHUTDOWN_THRESHOLD и SHUTDOWN_TIMEOUT должны быть неотрицательными!"
    temp_log = LOG_FILE if LOG_FILE else "ups_log.txt"
    with open(temp_log, 'a') as f:
        f.write(f"{datetime.now().strftime('%d.%m.%Y %H:%M:%S')} - {error_msg}\n")
    sys.exit(1)

# --- Функции ---

# Функция записи сообщений в лог-файл
# Создаёт новый заголовок в логе при смене дня, иначе добавляет сообщения
def write_to_log(message):
    today = datetime.now().strftime('%d.%m.%Y')
    last_reset_day = None

    # Проверка даты последней модификации лог-файла
    if os.path.exists(LOG_FILE):
        last_modified = datetime.fromtimestamp(os.path.getmtime(LOG_FILE)).strftime('%d.%m.%Y')
        last_reset_day = last_modified

    # Если день сменился, создаём новый заголовок в логе
    if today != last_reset_day:
        with open(LOG_FILE, 'w') as f:
            f.write(f"=== Лог за {today} ===\n")

    # Добавляем сообщение в лог с текущей датой и временем
    with open(LOG_FILE, 'a') as f:
        f.write(f"{datetime.now().strftime('%d.%m.%Y %H:%M:%S')} - {message}\n")

# Функция отправки уведомлений в Telegram
# Отправляет сообщение в указанный чат, логирует ошибки при сбое
def send_to_telegram(text):
    try:
        requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage",
            data={"chat_id": TELEGRAM_CHAT_ID, "text": text},
            timeout=5
        )
    except requests.exceptions.RequestException as e:
        write_to_log(f"Ошибка Telegram: {str(e)}")

# Функция получения состояния батареи через WMI
# Возвращает процент заряда, статус подключения и оставшееся время работы
def get_battery_status():
    try:
        wmi = win32com.client.GetObject("winmgmts:")
        batteries = wmi.InstancesOf("Win32_Battery")
        
        for battery in batteries:
            percent = battery.EstimatedChargeRemaining
            status = battery.BatteryStatus
            runtime = battery.EstimatedRunTime
            
            # Проверка доступности данных о заряде
            if percent is None:
                write_to_log("Ошибка: процент заряда батареи не обнаружен!")
                sys.exit(1)
            
            # plugged: True, если подключено к сети (BatteryStatus == 2)
            plugged = (status == 2)
            # runtime: время работы в минутах, 0 если данные недоступны
            runtime = runtime if runtime != 0xFFFFFFFE else 0
            
            return percent, plugged, runtime
        
        write_to_log("Ошибка: батарея не обнаружена!")
        sys.exit(1)
        
    except Exception as e:
        write_to_log(f"Ошибка WMI: {str(e)}")
        sys.exit(1)

# --- Основная логика ---
# Отслеживает состояние батареи, отправляет уведомления и выключает ПК при низком заряде
def main():
    power_lost_time = None      # Время отключения питания
    was_on_battery = False      # Флаг работы от батареи
    telegram_notified = False   # Флаг отправки уведомления в Telegram
    remaining_time_at_loss = 0  # Оставшееся время при отключении питания
    charge_at_loss = 0          # Заряд батареи при отключении питания

    while True:
        # Получение текущего состояния батареи
        charge, plugged, remaining_time = get_battery_status()

        # Обработка отключения питания
        if not plugged and not was_on_battery:
            was_on_battery = True
            power_lost_time = datetime.now()
            remaining_time_at_loss = remaining_time
            charge_at_loss = charge
            log_msg = f"Отключение электричества! Заряд: {charge_at_loss}%, осталось: {remaining_time_at_loss} мин."
            write_to_log(log_msg)
            # Уведомление отправляется в следующем блоке после задержки

        # Уведомление в Telegram после задержки
        if was_on_battery and not telegram_notified and not plugged:
            if (datetime.now() - power_lost_time).seconds >= DELAY_NOTIFY:
                send_to_telegram(
                    f"{power_lost_time.strftime('%d.%m.%Y %H:%M:%S')} - Отключение электричества!\n"
                    f"Заряд: {charge_at_loss}%, осталось: {remaining_time_at_loss} мин."
                )
                telegram_notified = True

        # Восстановление питания
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

        # Выключение при низком заряде
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

        # Пауза перед следующей проверкой
        time.sleep(CHECK_INTERVAL)

# Запуск скрипта с обработкой прерывания (Ctrl+C)
if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(0)