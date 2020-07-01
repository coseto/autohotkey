driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
driver.Get("http://192.168.50.15:9900/search?rangetype=relative&fields=facility%2Clevel%2Cgl2_message_id%2Cmessage%2CWebSocketResponseURL%2Cline%2CthreadName%2Ctimestamp%2Csource%2CloggerName%2Cfile&width=1522&highlightMessage=&relative=300&q=")
driver.findElementByid("username").SendKeys("admin")
driver.findElementByid("password").SendKeys("admin123456")
Xpath1 = //html/body/div/div/div/div/form/div[3]/button
driver.findElementByXPath(Xpath1).click()
/*
element := driver.findElementByName("q")	;查找name為 q 的元素
element.Clear()								;清除元素欄位資料
element.SendKeys("flutter")					;在元素欄位輸入文字"flutter"
element.SendKeys(driver.Keys.Control, "a")	;組合鍵ctrl+a 
element.SendKeys(driver.Keys.Control, "v")	;組合鍵ctrl+c
*/
Xpath3 = //html/body/div/div/div[2]/div[2]/div[2]/div/div/div[2]/div[6]/div[3]/div/div/table/tbody/tr/td[6]
driver.findElementByXPath(Xpath3).click()
Xpath4 = //html/body/div/div/div[2]/div[2]/div[2]/div/div/div[2]/div[6]/div[3]/div/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/dl/span[2]/dd/div[2]
xpath5 := driver.findElementByXPath(Xpath4).text
MsgBox, %xpath5%
