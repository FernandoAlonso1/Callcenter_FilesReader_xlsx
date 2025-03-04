package main

import (
	"fmt"
	"log"
	"time"

	"github.com/tealeg/xlsx"
)

// Функция для преобразования времени в секунды
func timeToSeconds(t time.Time) int {
	return t.Hour()*3600 + t.Minute()*60 + t.Second()
}

// Функция преобразования секунд в формат времени 00:00:00
func secondsToTime(seconds int) string {
	hours := seconds / 3600
	seconds %= 3600
	minutes := seconds / 60
	seconds %= 60
	return fmt.Sprintf("%02d:%02d:%02d", hours, minutes, seconds)
}

func main() {
	// Открываем файл input.xlsx входные данные номера/кол-во звонков/время звонков
	fileInput, err := xlsx.OpenFile("input.xlsx")
	if err != nil {
		log.Fatalf("Не удалось открыть файл input.xlsx: %s", err)
	}

	// Открываем файл filter.xlsx данные номера/города
	fileFilter, err := xlsx.OpenFile("filter.xlsx")
	if err != nil {
		log.Fatalf("Не удалось открыть файл filter.xlsx: %s", err)
	}

	//данные находятся в первом листе каждого файла
	sheetInput := fileInput.Sheets[0]
	sheetFilter := fileFilter.Sheets[0]

	// Создаем map для хранения сумм по A2 и A3, так же для связи B1, B2
	type sumData struct {
		sumA2 float64
		sumA3 int
		b2    string
	}
	sumMap := make(map[string]sumData)

	//map связи b1 b2
	b1ToB2 := make(map[string]string)

	// Проходим по строкам файла filter и создаем map для быстрого поиска
	for _, row := range sheetFilter.Rows {
		if len(row.Cells) < 2 {
			continue // Пропускаем строки, где меньше двух колонок
		}
		b1 := row.Cells[0].String() // B1
		b2 := row.Cells[1].String() // B2
		b1ToB2[b1] = b2
	}

	// Проходим по строкам файла input и суммируем значения A2 и A3 для совпавших A1 и B1
	for _, row := range sheetInput.Rows {
		if len(row.Cells) < 3 {
			continue // Пропускаем строки, где меньше трех колонок
		}
		a1 := row.Cells[0].String()     // A1
		a2, err := row.Cells[1].Float() // A2
		if err != nil {
			log.Printf("Ошибка преобразования A2 в число: %s", err)
			continue
		}
		a3 := row.Cells[2].String() // A3 время в формате 00:00:00
		if a3 == "" {
			a3 = "00:00:00"
		}

		t, err := time.Parse("15:04:05", a3)
		if err != nil {
			log.Printf("Ошибка преобразования A3 в число: %s", err)
			continue
		}
		seconds := timeToSeconds(t)

		// b2 по b1 (a1)
		b2, exist := b1ToB2[a1]
		if !exist {
			continue
		}

		data := sumMap[b2]
		data.sumA2 += a2
		data.sumA3 += seconds
		sumMap[b2] = data
	}

	// Создаем новый файл result.xlsx
	fileC := xlsx.NewFile()
	sheetC, err := fileC.AddSheet("Result")
	if err != nil {
		log.Fatalf("Не удалось создать лист для файла result.xlsx: %s", err)
	}

	// Записываем данные в файл result
	for b2, data := range sumMap {
		row := sheetC.AddRow()
		row.AddCell().SetString(b2)                        // B2
		row.AddCell().SetFloat(data.sumA2)                 // Sum A2
		row.AddCell().SetString(secondsToTime(data.sumA3)) // Sum A3 в формате 00:00:00
	}

	// Сохраняем файл result.xlsx
	err = fileC.Save("result.xlsx")
	if err != nil {
		log.Fatalf("Не удалось сохранить файл result.xlsx: %s", err)
	}

	fmt.Println("Результат успешно записан в файл result.xlsx")
}
