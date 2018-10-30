// certificateNotMortgageEmailWithDelearName project main.go
package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"runtime"
	"strconv"
	"strings"
	"time"

	"github.com/axgle/mahonia"
	"github.com/larspensjo/config"
	"github.com/smartwalle/going/email"
	"github.com/tealeg/xlsx"
)

var (
	configFile = flag.String("configfile", "Sender_config.ini", "General configuration file")
	Version    = "riskEmailSending V2.0.20170601 "
	Auther     = "Owen Shen"
)
var body, emailTitle, tmpTitle string
var dateRiqi []int

//var webtitleBar string
//var webSplitSlice string

func getArgs() {
	version := flag.Bool("v", false, "version")
	flag.Parse()
	if *version {
		fmt.Println("Version：", Version)
		fmt.Println("Auther:", Auther)
		return
	}
}

func main() {
	if len(os.Args) > 1 {
		getArgs()
	} else {
		//		TOPIC := readconfigfile()
		//		SendEmail(TOPIC)
		sendingEmail()
		//				FIEmailList, SMEmailList := readDealerEmailAddress("test112")
		//				fmt.Println(FIEmailList, SMEmailList)
	}
}

func sendingEmail() {
	TOPIC := readconfigfile()
	//详细内容文件名
	content := TOPIC["content"]
	fmt.Println("内容文件" + content)
	//邮件地址
	email := TOPIC["email"]
	fmt.Println("邮件地址文件" + email)
	fmt.Println("eselect", TOPIC["eselect"], "emailto", TOPIC["emailto"], "emailcc", TOPIC["emailcc"])
	fmt.Println("cselect", TOPIC["cselect"], "emailcontentlen", TOPIC["emailcontentlen"])
	ti := time.Now().Format("20060102")
	tii := time.Now().Format("20060102150405")
	os.IsExist(os.Mkdir(ti, os.ModePerm))
	//生成文件
	logFile, _ := os.OpenFile(ti+"/"+tii+".txt", os.O_RDWR|os.O_CREATE, 0666)
	defer logFile.Close()
	//读取经销商文件内容
	contentSlice, _ := xlsx.FileToSlice(content)

	rows := len(contentSlice[0])
	logFile.WriteString("内容有：" + strconv.Itoa(rows) + "行数据。\r\n")
	logFile.WriteString("发送邮箱为：" + TOPIC["username"] + "；密码：" + TOPIC["password"] + "\r\n")
	logFile.WriteString("如果程序出现错误请检查配置文件是否正确。\r\n从Sebnder_config.ini邮件配置，模版格式等，确认都没问题之后再联系IT\r\n***********～。～*******\r\n")
	//这里要加一个变量，以提取文件的title
	x, _ := strconv.Atoi(TOPIC["emailcontentlen"])
	titleBars := contentSlice[0][0][:x]
	//webtitleBar = titleBars //网页上显示用utf-8
	conv := mahonia.NewEncoder("GBK")
	//webconv := mahonia.NewEncoder("UTF-8")

	for i, v := range titleBars {
		titleBars[i] = conv.ConvertString(v)
		//fmt.Println(v)
		if strings.Contains(v, "日") {
			dateRiqi = append(dateRiqi, i)
		}
		//webtitleBar = webtitleBar + " | " + v
		//webtitleBar[i] = webconv.ConvertString(v)
		//fmt.Println("转成gbk:" + titleBars[i])
	}
	fmt.Println(dateRiqi)
	//emailTitle = TOPIC["emailTitle"]

	tmpTitle = ti + "广汇月供扣款失败客户名单"
	//fmt.Println(webtitleBar)
	//************这里也要加变量以确定依据的内容
	//*******以经销商名为依据合并内容发送邮件*********************************
	x, _ = strconv.Atoi(TOPIC["cselect"])
	x = x - 1
	DealerName := contentSlice[0][rows-1][x]
	//fmt.Println(DealerName)
	logFile.WriteString("使用内容表第" + TOPIC["cselect"] + "列，以及邮件表第" + TOPIC["eselect"] + "列匹配邮件地址\r\n")
	logFile.WriteString("匹配成功后Email表第" + TOPIC["emailto"] + "列，作为收件箱\r\n")
	if TOPIC["emailccto"] == "well" {
		logFile.WriteString("Email表第" + TOPIC["emailcc"] + "列，作为抄收件箱\r\n")
	} else if TOPIC["emailccto"] == "yes" {
		logFile.WriteString("合并Email表第" + TOPIC["emailcc"] + "列，作为收件箱\r\n")
	} else if TOPIC["emailccto"] == "no" {
		logFile.WriteString("Email表里不设置抄收邮箱。\r\n ")

	}
	SplitSlice := [][]string{}
	//拆分Excel，并回写内容
	for i := rows - 1; i >= 0; i-- {
		DealerName, SplitSlice = splitE(logFile, ti, DealerName, email, i, contentSlice, SplitSlice, TOPIC, titleBars)
	}
}

func readDealerEmailAddress(DealerName, email string, TOPIC map[string]string) (ToEmailList, CcEmailList [1]string) {
	//	DEAExcelFileName := "testingEamilAddress.xlsx"
	emailSlice, _ := xlsx.FileToSlice(email)
	//fmt.Print("打印email表：")
	//fmt.Println(emailSlice)
	row := len(emailSlice[0])
	//fmt.Println(row)
	x, _ := strconv.Atoi(TOPIC["eselect"])
	x = x - 1
	y, _ := strconv.Atoi(TOPIC["emailto"])
	y = y - 1
	z, _ := strconv.Atoi(TOPIC["emailcc"])
	z = z - 1
	for i := row - 1; i >= 0; i-- {
		//********************************以经销商名为依据生成邮件地址************************
		if DealerName == emailSlice[0][i][x] {
			if emailSlice[0][i][y] != "" {
				ToEmailList[0] = emailSlice[0][i][y]
			} else {
				ToEmailList[0] = ""
			}
			if TOPIC["emailccto"] == "yes" {
				if emailSlice[0][i][z] != "" {
					ToEmailList[0] = ToEmailList[0] + "," + emailSlice[0][i][z]
				}
			} else if TOPIC["emailccto"] == "no" {
				CcEmailList[0] = ""
			} else if TOPIC["emailccto"] == "well" {
				if emailSlice[0][i][z] != "" {
					CcEmailList[0] = emailSlice[0][i][z]
				} else {
					CcEmailList[0] = ""
				}
			}

		}
	}
	return ToEmailList, CcEmailList
}

func splitE(logFile *os.File, ti, DealerName, email string, i int, contentSlice [][][]string, SplitSlice [][]string, TOPIC map[string]string, titleBars []string) (dn string, SplitS [][]string) {
	//***************************************以经销商名为依据合并内容发送邮件*********************************
	x, _ := strconv.Atoi(TOPIC["cselect"])
	x = x - 1
	k, _ := strconv.Atoi(TOPIC["emailcontentlen"])
	//fmt.Println("k", k)
	if DealerName == contentSlice[0][i][x] {
		tmpArr := []string{}
		//***************************************要拼接的内容行数*********************************
		for j := 0; j < k; j++ {
			if j == 0 {
				tmp := "'" + contentSlice[0][i][j]
				tmpArr = append(tmpArr, tmp)
			} else {
				//fmt.Println(contentSlice[0][0][j])
				t := 0
				for _, v := range dateRiqi {

					if j == v {
						fmt.Println("修改日期")
						t1, _ := strconv.ParseInt(contentSlice[0][i][j], 10, 64)
						//fmt.Println(t)
						t1 = (t1-70*365-19)*60*60*24 - 8*3600
						fmt.Println(t)

						ti := time.Unix(t1, 0)
						date := ti.Format("2006/01/02")
						tmpArr = append(tmpArr, date)
						t = t + 1

						//						tmp := strings.Split(contentSlice[0][i][j], "/")
						//						if len(tmp) == 3 {
						//							date := "20" + tmp[2] + "年" + tmp[1] + "月" + tmp[0] + "日"
						//							tmpArr = append(tmpArr, date)
						//						}
						//						//tmpArr = append(tmpArr, contentSlice[0][i][j])
						//						t = t + 1
					}
					//tmpArr = append(tmpArr, contentSlice[0][i][j])

				}
				if t == 0 {
					//fmt.Println("hhhhhhhhh")
					tmpArr = append(tmpArr, contentSlice[0][i][j])
				}

			}

		}
		SplitSlice = append(SplitSlice, tmpArr)
		//		webSplitSlice = SplitSlice
		return DealerName, SplitSlice
	} else {
		fmt.Println(SplitSlice)
		ExportFileName := splitExcel(ti, i, contentSlice, SplitSlice, TOPIC, titleBars)

		ToEmailList, CcEmailList := readDealerEmailAddress(DealerName, email, TOPIC)
		//		fmt.Println(DealerName)
		if ToEmailList[0] == "" && CcEmailList[0] == "" {
			logFile.WriteString(DealerName + " no To email or Cc address.\r\n")
			//fmt.Println(DealerName + " no To email address." + "\r\n")
		} else {
			logFile.WriteString(DealerName + "; attachment is " + ExportFileName + "\r\n")
			//fmt.Println("输出body哦-----------")
			//conv1 := mahonia.NewEncoder("gbk")
			//			for j, v := range SplitSlice {
			//				if j < 15 {

			//					for i := 0; i < 15; i++ {
			//						//x := conv1.ConvertString(v[i])
			//						body = body + v[i] + "|"
			//					}
			//					body = body + "\r\n"
			//				}
			//			}
			//			var titleBar string
			//			for _, v := range webtitleBar {
			//				x := conv1.ConvertString(v)
			//				titleBar = titleBar + x + "  |  "
			//			}
			//titleBar = conv1.ConvertString(titleBar)
			body = DealerName + "\r\n"
			emailTitle = tmpTitle + "(" + DealerName + ")"
			//fmt.Println(body)
			SendEmail(TOPIC, ToEmailList, CcEmailList, ExportFileName, logFile)
		}

		SplitSlice = [][]string{}
		//emailTitle = ""
		//***************************************以经销商名为依据合并内容发送邮件*********************************
		DealerName = contentSlice[0][i][x]
		DealerName, SplitSlice = splitE(logFile, ti, DealerName, email, i, contentSlice, SplitSlice, TOPIC, titleBars)
		return DealerName, SplitSlice
	}
}

func splitExcel(ti string, i int, contentSlice [][][]string, SplitSlice [][]string, TOPIC map[string]string, titleBars []string) (ExportFileName string) {
	x, _ := strconv.Atoi(TOPIC["cselect"])
	x = x - 1
	ExportFile, _ := os.Create(ti + "/" + contentSlice[0][i+1][x] + ".csv")
	defer ExportFile.Close()
	WriteFile := csv.NewWriter(ExportFile)
	conv := mahonia.NewEncoder("GBK")
	//webconv := mahonia.NewEncoder("UTF-8")

	for i, w := range SplitSlice {
		for j, x := range w {
			SplitSlice[i][j] = conv.ConvertString(x)
			//			if i < 15 {
			//				webSplitSlice = webSplitSlice + "|" + x
			//			}

		}
		//		if i < 15 {
		//			webSplitSlice = webSplitSlice + "\r\n"
		//		}

	}
	WriteFile.Write(titleBars)
	WriteFile.WriteAll(SplitSlice)
	WriteFile.Flush()

	ExportFileName = ExportFile.Name()
	return ExportFileName
}

func SendEmail(TOPIC map[string]string, ToEmailList, CcEmailList [1]string, ExportFileName string, logFile *os.File) {

	var config = &email.MailConfig{}
	config.Username = TOPIC["username"]
	config.Host = TOPIC["host"]
	config.Password = TOPIC["password"]
	config.Port = TOPIC["port"]
	config.Secure = false

	//title = title + emailTitle
	var e = email.NewTextMessage(emailTitle, "")
	e.From = TOPIC["from"]
	//get current dealer's email address
	//e.To = []string{ToEmailList[0]}
	if ToEmailList[0] != "" {
		e.To = strings.Split(ToEmailList[0], ",")
		logFile.WriteString(" has sent To " + ToEmailList[0] + "\r\n")
	} else {
		logFile.WriteString("没有接受邮箱。")
	}

	//fmt.Println(ToEmailList[0])
	//get current dealer's cc email address
	if len(TOPIC["cc"]) != 0 {
		//e.Cc = []string{CcEmailList[0]}
		tmpemail := TOPIC["cc"]
		if CcEmailList[0] != "" {
			tmpemail = tmpemail + "," + CcEmailList[0]
		}
		e.Cc = strings.Split(tmpemail, ",")
		logFile.WriteString("has sent Cc  :" + tmpemail + "\r\n")
	} else {
		logFile.WriteString("请在配置文件里设置本人要抄送的邮箱。")
	}
	logFile.WriteString("-----------------------------\r\n")
	//	e.Cc = []string{SMEmailList[0]}
	//e.Bcc = []string{"dzamd@dongzhengafc.com"}
	//e.Bcc = []string{TOPIC["bcct"]}
	b, _ := ioutil.ReadFile(TOPIC["emailBody"])
	body = body + string(b)
	//e.Content = string(b)
	e.Content = body
	//get current dealer's email attachment
	if TOPIC["attach"] == "yes" {
		e.AttachFile(ExportFileName)
	}
	//e.AttachFile(ExportFileName)
	err := email.SendMail(config, e)
	prerr(err)
	body = " "
}

func readconfigfile() (TOPIC map[string]string) {
	TOPIC = make(map[string]string)
	runtime.GOMAXPROCS(runtime.NumCPU())
	flag.Parse()

	//set config file std
	cfg, err := config.ReadDefault(*configFile)
	if err != nil {
		log.Fatalf("Fail to find", *configFile, err)
	}
	//set config file std End

	//Initialized topic from the configuration
	if cfg.HasSection("topicArr") {
		section, err := cfg.SectionOptions("topicArr")
		if err == nil {
			for _, v := range section {
				options, err := cfg.String("topicArr", v)
				if err == nil {
					TOPIC[v] = options
				}
			}
		}
	}
	//Initialized topic from the configuration END
	return TOPIC
}

func prerr(err error) {
	if err != nil {
		panic(err)
		var stop string
		fmt.Scanln(&stop)
	}
}
