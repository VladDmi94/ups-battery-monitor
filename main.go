package main

import (
	"bufio"
	"fmt"
	"log"
	"net/http"
	"net/url"
	"os"
	"os/exec"
	"strconv"
	"strings"
	"time"

	"github.com/yusufpapurcu/wmi"
)

// Структура для чтения Win32_Battery через WMI
type Win32_Battery struct {
	EstimatedChargeRemaining uint16
	BatteryStatus            uint16
	EstimatedRunTime         uint32
}

// Конфигурация приложения
type Config struct {
	TelegramToken     string
	TelegramChatID    string
	CheckInterval     time.Duration
	DelayNotify       float64
	ShutdownThreshold uint16
	ShutdownTimeout   int
	LogFile           string
}

// Простой парсер .env файла без сторонних библиотек
func loadEnv(filePath string) map[string]string {
	env := make(map[string]string)
	file, err := os.Open(filePath)
	if err != nil {
		return env
	}
	defer file.Close()

	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		line := strings.TrimSpace(scanner.Text())
		if line == "" || strings.HasPrefix(line, "#") {
			continue
		}
		parts := strings.SplitN(line, "=", 2)
		if len(parts) == 2 {
			env[strings.TrimSpace(parts[0])] = strings.TrimSpace(parts[1])
		}
	}

	// Проверка на ошибки считывания файла после цикла
	if err := scanner.Err(); err != nil {
		log.Printf("Ошибка чтения .env файла: %v", err)
	}

	return env
}

// Функция записи в лог-файл и вывода в консоль
func writeLog(logFile, message string) {
	timestamp := time.Now().Format("02.01.2006 15:04:05")
	entry := fmt.Sprintf("%s - %s\n", timestamp, message)

	fmt.Print(entry) // Вывод в консоль

	f, err := os.OpenFile(logFile, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
	if err != nil {
		log.Printf("Ошибка записи в лог: %v", err)
		return
	}
	defer f.Close()
	f.WriteString(entry)
}

// Отправка сообщения в Telegram
func sendTelegram(token, chatID, text string) bool {
	apiURL := fmt.Sprintf("https://api.telegram.org/bot%s/sendMessage", token)
	data := url.Values{}
	data.Set("chat_id", chatID)
	data.Set("text", text)

	client := http.Client{Timeout: 5 * time.Second}
	resp, err := client.PostForm(apiURL, data)
	if err != nil {
		return false
	}
	defer resp.Body.Close()

	return resp.StatusCode == http.StatusOK
}

// Запрос статуса батареи через WMI
func getBatteryStatus() (uint16, bool, uint32, error) {
	var batteries []Win32_Battery
	query := wmi.CreateQuery(&batteries, "")
	err := wmi.Query(query, &batteries)
	if err != nil || len(batteries) == 0 {
		return 0, false, 0, fmt.Errorf("батарея не найдена или ошибка WMI: %v", err)
	}

	b := batteries[0]
	plugged := (b.BatteryStatus == 2)
	runtime := b.EstimatedRunTime
	if runtime == 0xFFFFFFFE {
		runtime = 0
	}

	return b.EstimatedChargeRemaining, plugged, runtime, nil
}

func main() {
	env := loadEnv(".env")

	logFile := env["LOG_FILE"]
	if logFile == "" {
		logFile = "ups_log.txt"
	}

	// Чтение параметров с дефолтными значениями
	token := env["TELEGRAM_TOKEN"]
	chatID := env["TELEGRAM_CHAT_ID"]

	if token == "" || chatID == "" {
		writeLog(logFile, "Ошибка: TELEGRAM_TOKEN или TELEGRAM_CHAT_ID не заданы в .env!")
		os.Exit(1)
	}

	checkIntervalSec, _ := strconv.Atoi(env["CHECK_INTERVAL"])
	if checkIntervalSec <= 0 {
		checkIntervalSec = 5
	}

	delayNotifySec, _ := strconv.ParseFloat(env["DELAY_NOTIFY"], 64)
	if delayNotifySec < 0 {
		delayNotifySec = 60
	}

	shutdownThreshold, _ := strconv.Atoi(env["SHUTDOWN_THRESHOLD"])
	if shutdownThreshold <= 0 {
		shutdownThreshold = 20
	}

	shutdownTimeout, _ := strconv.Atoi(env["SHUTDOWN_TIMEOUT"])
	if shutdownTimeout <= 0 {
		shutdownTimeout = 30
	}

	cfg := Config{
		TelegramToken:     token,
		TelegramChatID:    chatID,
		CheckInterval:     time.Duration(checkIntervalSec) * time.Second,
		DelayNotify:       delayNotifySec,
		ShutdownThreshold: uint16(shutdownThreshold),
		ShutdownTimeout:   shutdownTimeout,
		LogFile:           logFile,
	}

	writeLog(cfg.LogFile, "Мониторинг UPS (Go) успешно запущен.")

	var (
		wasOnBattery        bool
		telegramNotified    bool
		powerLostTime       time.Time
		chargeAtLoss        uint16
		remainingTimeAtLoss uint32
	)

	for {
		charge, plugged, remainingTime, err := getBatteryStatus()
		if err != nil {
			writeLog(cfg.LogFile, err.Error())
			time.Sleep(cfg.CheckInterval)
			continue
		}

		// 1. Фиксация отключения питания
		if !plugged && !wasOnBattery {
			wasOnBattery = true
			powerLostTime = time.Now()
			chargeAtLoss = charge
			remainingTimeAtLoss = remainingTime
			msg := fmt.Sprintf("Отключение электричества! Заряд: %d%%, осталось: %d мин.", chargeAtLoss, remainingTimeAtLoss)
			writeLog(cfg.LogFile, msg)
		}

		// 2. Уведомление в Telegram с задержкой
		if wasOnBattery && !telegramNotified && !plugged {
			elapsed := time.Since(powerLostTime).Seconds()
			if elapsed >= cfg.DelayNotify {
				msg := fmt.Sprintf("%s — Отключение электричества!\nЗаряд: %d%%, осталось: %d мин.",
					powerLostTime.Format("02.01.2006 15:04:05"), chargeAtLoss, remainingTimeAtLoss)
				if sendTelegram(cfg.TelegramToken, cfg.TelegramChatID, msg) {
					telegramNotified = true
				}
			}
		}

		// 3. Восстановление питания
		if plugged && wasOnBattery {
			duration := int(time.Since(powerLostTime).Seconds())
			msg := fmt.Sprintf("Электричество восстановлено. Заряд: %d%%, прошло: %d мин %d сек.",
				charge, duration/60, duration%60)
			writeLog(cfg.LogFile, msg)

			if float64(duration) >= cfg.DelayNotify {
				sendTelegram(cfg.TelegramToken, cfg.TelegramChatID,
					fmt.Sprintf("%s — Электричество восстановлено.\nЗаряд: %d%%, прошло: %d мин %d сек.",
						time.Now().Format("02.01.2006 15:04:05"), charge, duration/60, duration%60))
			}

			// Сброс флагов
			wasOnBattery = false
			telegramNotified = false
		}

		// 4. Выключение ПК при критическом заряде
		if wasOnBattery && charge > 0 && charge <= cfg.ShutdownThreshold {
			duration := int(time.Since(powerLostTime).Seconds())
			msg := fmt.Sprintf("Критический уровень заряда (%d%%)! Инициировано выключение ПК.", charge)
			writeLog(cfg.LogFile, msg)

			sendTelegram(cfg.TelegramToken, cfg.TelegramChatID,
				fmt.Sprintf("%s — Выключение ПК!\nЗаряд: %d%%, прошло: %d мин %d сек.",
					time.Now().Format("02.01.2006 15:04:05"), charge, duration/60, duration%60))

			// Выполнение команды shutdown в Windows
			cmd := exec.Command("shutdown", "/s", "/t", strconv.Itoa(cfg.ShutdownTimeout))
			cmd.Run()
			os.Exit(0)
		}

		time.Sleep(cfg.CheckInterval)
	}
}
