package main

import "fmt"

func main() {
	var a = []int{1, 2, 3, 4, 5, 6, 7, 8, 9, 10}
	var b = []int{11, 12, 13, 14, 15, 16, 17, 18, 19, 20}

	a = append(a, b...)
	fmt.Println(c)
}
