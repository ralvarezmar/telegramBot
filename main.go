package main

import (
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
	tgbotapi "github.com/go-telegram-bot-api/telegram-bot-api"
)

func main() {
	token := "656177777:AAFMuW8U8e1ob_gh4O4aFJJZzyjTSZxJyvo"
	bot, err := tgbotapi.NewBotAPI(token)
	if err != nil {
		log.Panic(err)
	}
	//home := os.Getenv("HOME")
	home := ".."
	path := home + "/bbdd.xlsx"
	var f *excelize.File

	if _, err := os.Stat(path); os.IsNotExist(err) {
		f = excelize.NewFile()
		log.Print("File created")
		if err := f.SaveAs(path); err != nil {
			log.Fatal(err)
		}
	}
	bot.Debug = true

	log.Printf("Authorized on account %s", bot.Self.UserName)

	u := tgbotapi.NewUpdate(0)
	u.Timeout = 60

	updates, err := bot.GetUpdatesChan(u)

	for update := range updates {
		value, newDate, f := readExcel(path)
		if update.Message == nil {
			continue
		}
		msgReceived := update.Message.Text
		log.Printf("[%s] %s", update.Message.From.UserName, update.Message.Text)
		if update.Message.From.UserName != "Ramii5" {
			continue
		}
		if newDate {
			log.Println("nueva fecha")
			newRow(value, f)
		}
		column := getColumn(value, f)
		number, err := strconv.Atoi(msgReceived)
		if err != nil {
			msg := tgbotapi.NewMessage(update.Message.Chat.ID, "Manda un número, abollao")
			msg.ReplyToMessageID = update.Message.MessageID
			bot.Send(msg)
			log.Println(err)
			continue
		} else {
			f.SetCellValue("Sheet1", column, number)
			msg := tgbotapi.NewMessage(update.Message.Chat.ID, "Apuntado")
			msg.ReplyToMessageID = update.Message.MessageID
			bot.Send(msg)
		}
		if err := f.SaveAs(path); err != nil {
			log.Fatal(err)
		}
	}
}

func newRow(value int, f *excelize.File) {
	date := time.Now().Format("02 Jan 06 ")
	cell := fmt.Sprintf("A%d", value)
	f.SetCellValue("Sheet1", cell, date)
}

func getColumn(value int, f *excelize.File) string {
	for i := 1; i < 25; i++ {
		letter := toCharStr(i)
		row := fmt.Sprintf("%s%d", letter, value)
		cellVal := f.GetCellValue("Sheet1", row)
		if cellVal == "" {
			return row
		}
	}
	return ""
}

func toCharStr(i int) string {
	return string('A' - 1 + i)
}

func readExcel(path string) (value int, newDate bool, f *excelize.File) {
	f, _ = excelize.OpenFile(path)
	value = 1
	newDate = false
	for true {
		row := fmt.Sprintf("A%d", value)
		cellVal := f.GetCellValue("Sheet1", row)

		if strings.Contains(cellVal, time.Now().Format("02 Jan 06")) {
			newDate = false
			return
		}
		if cellVal == "" {
			newDate = true
			return
		}
		value++
	}
	return
}
