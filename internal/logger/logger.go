package logger

import (
	"log/slog"
	"os"
	"path/filepath"
)

var Logger *slog.Logger

func init() {
	// Create logs directory
	os.MkdirAll("logs", 0755)
	
	// Create log file
	logFile, err := os.OpenFile(filepath.Join("logs", "sheetfmt.log"), os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
	if err != nil {
		panic(err)
	}

	// Create structured logger that writes to both file and stdout
	Logger = slog.New(slog.NewTextHandler(logFile, &slog.HandlerOptions{
		Level: slog.LevelInfo,
	}))
}

func Info(msg string, args ...any) {
	Logger.Info(msg, args...)
}

func Error(msg string, args ...any) {
	Logger.Error(msg, args...)
}

func Debug(msg string, args ...any) {
	Logger.Debug(msg, args...)
}

func Warn(msg string, args ...any) {
	Logger.Warn(msg, args...)
}