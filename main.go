package main

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io/fs"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"
)

const (
	iDate  = 1
	iFrom1 = 3
	iTo1   = 4
	iFrom2 = 5
	iTo2   = 6
)

const (
	lScheduleStart = 9
)

const debug = false

func main() {
	var err error
	err = os.RemoveAll("/home/tikinang/documents/gabca_evidence_clean")
	if err != nil {
		panic(err)
	}

	err = os.MkdirAll("/home/tikinang/documents/gabca_evidence_clean", 0755)
	if err != nil {
		panic(err)
	}

	err = filepath.WalkDir("/home/tikinang/documents/gabca_evidence", func(path string, d fs.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if d.IsDir() {
			return nil
		}
		if !strings.HasSuffix(d.Name(), ".xlsx") {
			return nil
		}
		fmt.Println("processing:", d.Name())
		f, err := excelize.OpenFile(path)
		if err != nil {
			return err
		}
		defer f.Close()

		rows, err := f.GetRows("VZOR")
		if err != nil {
			return err
		}

		var weeks []*WeekSchedule
		var week *WeekSchedule
		for i, row := range rows {
			if i < lScheduleStart {
				continue
			}

			if debug {
				fmt.Println()
			}

			if len(row) < 4 {
				if debug {
					fmt.Println("skipping, short row:", i+1)
				}
				continue
			}

			if debug {
				fmt.Println("row number:", i+1)
				printRow(row)
			}

			if isEmpty(row[iFrom1]) && isEmpty(row[iTo1]) && isEmpty(row[iFrom2]) && isEmpty(row[iTo2]) {
				if debug {
					fmt.Println("skipping, no schedule:", i+1)
				}
				continue
			}

			schedule, err := getDaySchedule(row)
			if err != nil {
				fmt.Println("error getting day schedule:", err)
				printRow(row)
				continue
			}
			if debug {
				fmt.Println(schedule)
			}

			if schedule.Weekday == time.Monday {
				if week != nil {
					weeks = append(weeks, week)
				}
				week = &WeekSchedule{
					Days: make([]*DaySchedule, 7),
				}
			}
			if week != nil {
				week.Days[schedule.Weekday] = schedule
			}
		}

		if week != nil {
			weeks = append(weeks, week)
		}

		var prev *string
		for i, w := range weeks {
			if debug {
				fmt.Println(w)
			}
			this := fmt.Sprint(w)
			if prev != nil && *prev != this && i != len(weeks)-1 {
				fmt.Printf("MISMATCH!\n%s\n%s\n", *prev, this)
			}
			prev = &this
		}

		if len(weeks) == 0 {
			fmt.Println("NO SCHEDULE!")
			return nil
		}

		name, err := f.GetCellValue("Zaměstnanec", "B1")
		if err != nil {
			return err
		}
		surname, err := f.GetCellValue("Zaměstnanec", "B2")
		if err != nil {
			return err
		}
		position, err := f.GetCellValue("Zaměstnanec", "B3")
		if err != nil {
			return err
		}

		info := &Info{
			Schedule: weeks[1],
			Filename: d.Name(),
			Worker:   fmt.Sprintf("%s %s", name, surname),
			Position: strings.Join(strings.Fields(position), " "),
		}

		// FIXME(tikinang): Year config. Debug config.
		err = writeExcel(2024, info)
		if err != nil {
			return err
		}

		return nil
	})
	if err != nil {
		panic(err)
	}
}

type WeekSchedule struct {
	Days []*DaySchedule
}

type DaySchedule struct {
	Parts   []FromTo
	Weekday time.Weekday
}

func (d *DaySchedule) Hours() time.Time {
	var t time.Time
	for _, p := range d.Parts {
		t = t.Add(p.To.Sub(p.From))
	}
	return t
}

func (d *DaySchedule) String() string {
	fromTos := new(strings.Builder)
	for _, x := range d.Parts {
		fromTos.WriteString(fmt.Sprintf("%s => %s;", x.From.Format(time.TimeOnly), x.To.Format(time.TimeOnly)))
	}
	return fmt.Sprintf("%s: %s", d.Weekday, fromTos.String())
}

type FromTo struct {
	From time.Time
	To   time.Time
}

type Info struct {
	Schedule *WeekSchedule
	Filename string
	Worker   string
	Position string
}

func printRow(row []string) {
	for j, cell := range row {
		fmt.Println(j, cell)
	}
}

func parseTime(val string) (time.Time, error) {
	parts := strings.Split(val, ".")
	var clock string
	var leftover time.Duration
	if len(parts) == 1 {
		clock = parts[0]
	} else if len(parts) == 2 {
		clock = parts[0]
		a, err := strconv.ParseFloat("0."+parts[1], 64)
		if err != nil {
			return time.Time{}, err
		}
		leftover = time.Duration(math.Round(a))
	} else {
		return time.Time{}, fmt.Errorf("parts len invalid: %d, for %s", len(parts), val)
	}
	t, err := time.Parse(time.TimeOnly, clock)
	if err != nil {
		return time.Time{}, err
	}
	return t.Add(leftover * time.Second), nil
}

func isEmpty[T comparable](v T) bool {
	var empty T
	return empty == v
}

func getDaySchedule(row []string) (*DaySchedule, error) {
	date, err := time.Parse(time.DateOnly, row[iDate])
	if err != nil {
		return nil, err
	}
	s := &DaySchedule{
		Weekday: date.Weekday(),
	}

	if !isEmpty(row[iFrom1]) {
		from, err := parseTime(row[iFrom1])
		if err != nil {
			return nil, err
		}
		if !isEmpty(row[iTo1]) {
			to, err := parseTime(row[iTo1])
			if err != nil {
				return nil, err
			}
			s.Parts = append(s.Parts, FromTo{
				From: from,
				To:   to,
			})
		} else if !isEmpty(row[iTo2]) {
			to, err := parseTime(row[iTo2])
			if err != nil {
				return nil, err
			}
			s.Parts = append(s.Parts, FromTo{
				From: from,
				To:   to,
			})
		} else {
			return nil, errors.New("unexpected end of schedule, no end")
		}
	} else {
		return nil, errors.New("unexpected end of schedule, no from")
	}

	if !isEmpty(row[iFrom2]) {
		from, err := parseTime(row[iFrom2])
		if err != nil {
			return nil, err
		}
		if !isEmpty(row[iTo2]) {
			to, err := parseTime(row[iTo2])
			if err != nil {
				return nil, err
			}
			s.Parts = append(s.Parts, FromTo{
				From: from,
				To:   to,
			})
		} else {
			return nil, errors.New("unexpected end of schedule, no end to second from")
		}
	}

	return s, nil
}

var prague = must(time.LoadLocation("Europe/Prague"))

const (
	writeRowOffset = 6
	alphabet       = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
)

var weekdays = [7]string{
	"neděle",
	"pondělí",
	"úterý",
	"středa",
	"čtvrtek",
	"pátek",
	"sobota",
}

var months = [12]string{
	"leden",
	"únor",
	"březen",
	"duben",
	"květen",
	"červen",
	"červenec",
	"srpen",
	"září",
	"říjen",
	"listopad",
	"prosinec",
}

func writeExcel(year int, info *Info) error {
	f := excelize.NewFile()
	defer f.Close()

	var (
		cursor time.Month
		t      time.Time
	)

	sheet := func(cursor time.Month) string {
		return months[cursor-1]
	}

	cell := func(column string, value any) error {
		return f.SetCellValue(sheet(cursor), fmt.Sprintf("%s%d", column, t.Day()+writeRowOffset), value)
	}

	cellInfo := func(row int, values ...any) error {
		var err error
		for i, v := range values {
			err = f.SetCellValue(sheet(cursor), fmt.Sprintf("%s%d", string(alphabet[i]), row), v)
			if err != nil {
				return err
			}
		}
		return nil
	}

	var err error
	for t = time.Date(year, 1, 1, 0, 0, 0, 0, prague); t.Year() == year; t = t.AddDate(0, 0, 1) {
		if t.Month() != cursor {
			cursor = t.Month()

			_, err = f.NewSheet(sheet(cursor))
			if err != nil {
				return err
			}
		}

		// info

		err = cellInfo(1, "rok:", strconv.Itoa(year))
		if err != nil {
			return err
		}
		err = cellInfo(2, "měsíc:", strconv.Itoa(int(cursor)))
		if err != nil {
			return err
		}
		err = cellInfo(3, "jméno:", info.Worker)
		if err != nil {
			return err
		}
		err = cellInfo(4, "pozice a číslo:", info.Position)
		if err != nil {
			return err
		}
		err = cellInfo(6, "datum", "den", "příchod", "odchod", "příchod", "odchod", "hodin", "poznámka")
		if err != nil {
			return err
		}

		// days

		err = cell("A", t.Format(time.DateOnly))
		if err != nil {
			return err
		}
		err = cell("B", weekdays[t.Weekday()])
		if err != nil {
			return err
		}
		if day := info.Schedule.Days[t.Weekday()]; day != nil {
			if len(day.Parts) == 0 || len(day.Parts) > 2 {
				return fmt.Errorf("unexpected schedule parts: %+v", day.Parts)
			}
			if len(day.Parts) > 0 {
				schedule := day.Parts[0]
				err = cell("C", schedule.From.Format(time.TimeOnly))
				if err != nil {
					return err
				}
				err = cell("D", schedule.To.Format(time.TimeOnly))
				if err != nil {
					return err
				}
			}
			if len(day.Parts) == 2 {
				schedule := day.Parts[1]
				err = cell("E", schedule.From.Format(time.TimeOnly))
				if err != nil {
					return err
				}
				err = cell("F", schedule.To.Format(time.TimeOnly))
				if err != nil {
					return err
				}
			}
			err = cell("G", day.Hours().Format(time.TimeOnly))
			if err != nil {
				return err
			}
		}

		err = f.SetColWidth(sheet(cursor), "A", "A", 12)
		if err != nil {
			return err
		}
		err = f.SetColWidth(sheet(cursor), "B", "B", 9)
		if err != nil {
			return err
		}
		err = f.SetColWidth(sheet(cursor), "C", "H", 9)
		if err != nil {
			return err
		}
	}

	err = f.DeleteSheet("Sheet1")
	if err != nil {
		return err
	}

	err = f.SaveAs(fmt.Sprintf("/home/tikinang/documents/gabca_evidence_clean/%s", info.Filename))
	if err != nil {
		return err
	}

	return nil
}

func must[T any](val T, err error) T {
	if err != nil {
		panic(err)
	}
	return val
}
