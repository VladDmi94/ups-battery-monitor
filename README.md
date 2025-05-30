# UPS Battery Monitor

Скрипт для мониторинга батареи ИБП или ноутбука на Windows. Отслеживает отключение/восстановление питания, отправляет уведомления в Telegram и выключает ПК при низком заряде. Логи сохраняются в файл.

## Описание

- Отслеживает батарею (заряд, подключение, время) через WMI (`win32com.client`).
- Логирует события в `LOG_FILE`, указанный в `config.ini`.
- Уведомления в Telegram:
  - Отключение питания.
  - Восстановление.
  - Низкий заряд (с выключением).
- Работает фоново (например, через Планировщик задач).

Пример Telegram-уведомления:
```
30.05.2025 16:23:45 - Отключение электричества!
Заряд: 98%, осталось: 10 мин.
```

Пример лога:
```
=== Лог за 30.05.2025 ===
30.05.2025 16:23:45 - Отключение электричества! Заряд: 98%, осталось: 10 мин.
```

## Требования

- **ОС**: Windows 10/11 (64-бит).
- **Python**: Портативная версия в релизе.
- **Telegram**: Бот ([@BotFather](https://t.me/BotFather)), ID чата (@userinfobot).
- **Права**: Запись в лог, `shutdown`.

## Установка

### Вариант 1: Портативная версия (рекомендуется)
1. Скачайте `PythonEmbedded.zip` из [Releases](https://github.com/your-username/your-repo/releases), распакуйте в `C:\UPSMonitor`.
2. Создайте `config.ini` в `C:\UPSMonitor`:
   ```ini
   [Settings]
   TELEGRAM_TOKEN = ваш_токен
   TELEGRAM_CHAT_ID = ваш_чат_id
   CHECK_INTERVAL = 1
   DELAY_NOTIFY = 5
   LOG_FILE = ups_log.txt
   SHUTDOWN_THRESHOLD = 50
   SHUTDOWN_TIMEOUT = 3
   ```
3. Скопируйте `ups_monitor_wmi.py` в `C:\UPSMonitor`.
4. Запустите:
   ```bash
   C:\UPSMonitor\PythonEmbedded\python.exe C:\UPSMonitor\ups_monitor_wmi.py
   ```
   Или через `run.bat`.
5. Для фона: в Планировщике задач настройте запуск по событию `Microsoft-Windows-Kernel-Power`.

### Вариант 2: Python вручную
1. Установите Python 3.6+ ([python.org](https://www.python.org/downloads/)), добавьте в PATH.
2. Установите зависимости:
   ```bash
   pip install pywin32 requests
   ```
3. Создайте `config.ini`, запустите:
   ```bash
   python ups_monitor_wmi.py
   ```

## Конфигурация
В `config.ini` (`[Settings]`):
- `TELEGRAM_TOKEN`: токен бота.
- `TELEGRAM_CHAT_ID`: ID чата.
- `CHECK_INTERVAL`: интервал (сек, >0).
- `DELAY_NOTIFY`: задержка (сек, ≥0).
- `LOG_FILE`: лог (например, `ups_log.txt`).
- `SHUTDOWN_THRESHOLD`: порог заряда (%, ≥0).
- `SHUTDOWN_TIMEOUT`: время выключения (сек, ≥0).

Ошибки конфигурации пишутся в `LOG_FILE` или `ups_log.txt`.

## Логирование
Логи пишутся в `LOG_FILE`:
```
=== Лог за 30.05.2025 ===
30.05.2025 16:23:45 - Отключение электричества! Заряд: 98%, осталось: 10 мин.
```
Ошибки конфигурации — в `LOG_FILE` или `ups_log.txt`.

## Тестирование
1. Запустите скрипт.
2. Проверьте логи, Telegram, ошибки в `ups_log.txt` (без `config.ini`).
3. Тест выключения: `SHUTDOWN_THRESHOLD=99`, `SHUTDOWN_TIMEOUT=5`.

## Зависимости
- `pywin32` (MIT)
- `requests` (Apache 2.0)
- Зависимости: `urllib3` (MIT), `certifi` (MPL), `charset_normalizer` (MIT), `idna` (BSD)

## Замечания
- Только Windows (WMI, `shutdown`).
- `EstimatedRunTime` некорректен — 0.
- Интернет для Telegram.
- 64-бит Python Embedded (для 32-бит — другой).

## Лицензия
MIT (см. `LICENSE`).
