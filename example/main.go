package main

import (
	"fmt"
	"github.com/gofiber/fiber/v2"
	"github.com/tss182/excel"
)

type (
	Data struct {
		ID      int    `excel:"ID" json:"id"`
		Name    string `excel:"Name" json:"name"`
		Phone   int    `excel:"Phone Number" json:"phone"`
		Address string `excel:"Address" json:"address"`
	}
)

func main() {
	app := fiber.New()
	app.Post("/excel", func(c *fiber.Ctx) error {
		f, err := c.FormFile("file")
		if err != nil {
			return fmt.Errorf("failed to get file: %w", err)
		}

		fileReader, err := f.Open()

		fe, err := excel.OpenReader[Data](fileReader)
		if err != nil {
			return fmt.Errorf("failed to open reader: %w", err)
		}

		defer fe.Close()
		var limit uint = 10

		var data = make([]Data, 0, 1_000_000)

		err = fe.Read(&data, "Sheet1", excel.Opt{
			HeaderRow:    1,
			DataStartRow: 2,
			Limit:        limit,
		})
		if err != nil {
			return fmt.Errorf("failed to read: %w", err)
		}

		defer fe.CloseRow()

		if !fe.IsNext {
			return c.JSON(data)
		}

		for {
			var temp = make([]Data, 0, limit)
			err = fe.Next(&temp)
			if err != nil {
				return fmt.Errorf("failed to read next: %w", err)
			}
			data = append(data, temp...)
			if !fe.IsNext {
				break
			}
		}
		return c.JSON(data)
	})

	app.Listen(":8888")
}
