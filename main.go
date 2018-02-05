package main

import (
	"bufio"
	"fmt"
	"github.com/mengxiangjian13/docx"
	"os"
	"path/filepath"
	"regexp"
	"strings"
)

func main() {
	reader := bufio.NewReader(os.Stdin)
	fmt.Print("将word文档文件夹拖入此处，按回车结束：")
	text, _ := reader.ReadString('\n')
	path := strings.Replace(text, "\n", "", -1)
	fmt.Println("dir path :" + path)
	fmt.Print("输入搜索关键词，按回车结束：")
	keyword, _ := reader.ReadString('\n')
	keyword = strings.Replace(keyword, "\n", "", -1)
	fmt.Println("keyword :" + keyword)
	findFileByKeyword(path, keyword)
}

func findFileByKeyword(dirPath string, keyword string) {
	dir, err := os.Stat(dirPath)
	if err != nil {
		fmt.Println("没有找到该文件夹地址")
		return
	}
	if dir.IsDir() == false {
		fmt.Println("输入地址不是文件夹地址")
	}
	file, fileErr := os.Create(filepath.Join(dirPath, "result.txt"))
	defer file.Close()
	if fileErr != nil {
		fmt.Println("未知失败")
		return
	}
	result := ""
	walkFn := func(path string, info os.FileInfo, err error) error {
		if !info.IsDir() {
			ext := filepath.Ext(path)
			if ext == ".docx" {
				if isWordExist(path, keyword) {
					result = info.Name() + "\n"
				}
			}
		}
		return nil
	}
	filepath.Walk(dirPath, walkFn)
	file.Write([]byte(result))
}

func isWordExist(path string, keyword string) bool {
	docReplace, err := docx.ReadDocxFile(path)
	defer docReplace.Close()
	if err == nil {
		docx := docReplace.Editable()
		content := docx.TotalContent()
		r, _ := regexp.Compile("<w:t>[\\s\\S]*?</w:t>")
		rs := r.FindAllString(content, -1)
		exist := false
		for i := 0; i < len(rs); i++ {
			v := rs[i]
			temp := strings.Replace(v, "<w:t>", "", -1)
			temp = strings.Replace(temp, "</w:t>", "", -1)
			if strings.Contains(temp, keyword) {
				exist = true
				break
			}
		}
		return exist
	} else {
		fmt.Print(err)
		return false
	}
}
