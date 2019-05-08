using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace HistoricalDataCapture
{
	class Program
	{
		static void Main(string[] args)
		{
			//all strings
			string destinationFile = @"C:\Users\xadams\Desktop\TestData.txt";
			string TickerSymbol = "AAPL";
			
			
			//Itializes the chromedriver
			IWebDriver driver = new ChromeDriver(@"C:\Users\xadams\source\repos\Data-Capture-Auto\packages\Selenium.Chrome.WebDriver.73.0.0\driver");

			//Trys to do webpage automation
			try
			{
				driver.Manage().Window.Maximize();
			
				driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(15);

				driver.Navigate().GoToUrl($"https://finance.yahoo.com/quote/{TickerSymbol}/history?p={TickerSymbol}");			

				IWebElement date = driver.FindElement(By.CssSelector(@"#Col1-1-HistoricalDataTable-Proxy > section > div.Mt\28 15px\29.drop-down-selector.historical > div.Bgc\28 \24 extraLightGray\29.Bdrs\28 3px\29.P\28 10px\29 > div:nth-child(1) > span.Pos\28 r\29 > span > input"));
				date.Click();

				IWebElement maxYrs = driver.FindElement(By.CssSelector(@"#Col1-1-HistoricalDataTable-Proxy > section > div.Mt\28 15px\29.drop-down-selector.historical > div.Bgc\28 \24 extraLightGray\29.Bdrs\28 3px\29.P\28 10px\29 > div:nth-child(1) > span.Pos\28 r\29 > div > div.Ta\28 c\29.C\28 \24 gray\29 > span:nth-child(8)"));
				maxYrs.Click();

				IWebElement Done = driver.FindElement(By.CssSelector(@"#Col1-1-HistoricalDataTable-Proxy > section > div.Mt\28 15px\29.drop-down-selector.historical > div.Bgc\28 \24 extraLightGray\29.Bdrs\28 3px\29.P\28 10px\29 > div:nth-child(1) > span.Pos\28 r\29 > div > div.Mt\28 20px\29 > button.Bgc\28 \24 c-fuji-blue-1-b\29.Bdrs\28 3px\29.Px\28 20px\29.Miw\28 100px\29.Whs\28 nw\29.Fz\28 s\29.Fw\28 500\29.C\28 white\29.Bgc\28 \24 actionBlueHover\29 \3a h.Bd\28 0\29.D\28 ib\29.Cur\28 p\29.Td\28 n\29.Py\28 9px\29.Miw\28 80px\29 \21.Fl\28 start\29"));
				Done.Click();

				IWebElement Apply = driver.FindElement(By.CssSelector(@"#Col1-1-HistoricalDataTable-Proxy > section > div.Mt\28 15px\29.drop-down-selector.historical > div.Bgc\28 \24 extraLightGray\29.Bdrs\28 3px\29.P\28 10px\29 > button"));
				Apply.Click();

				int i;
				for (i = 0; i < 20; i++)
				{
					//scrolls down webpage to load more data on page before grabbing table
					IWebElement webPage = driver.FindElement(By.CssSelector(@"#Col1-1-HistoricalDataTable-Proxy > section > div.Pb\28 10px\29.Ovx\28 a\29.W\28 100\25 \29 > table > tfoot > tr > td"));
					((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", webPage);

					Thread.Sleep(2000);
				}
				
				//Finds data table with css selector
				IWebElement Table = driver.FindElement(By.CssSelector(@"#Col1-1-HistoricalDataTable-Proxy > section > div.Pb\28 10px\29.Ovx\28 a\29.W\28 100\25 \29 > table"));
				var TableRow = Table.FindElements(By.TagName("tr"));
				List<string> TextData = new List<string>();

				//For each row on table append it to csv file
				foreach (var Row in TableRow)
				{
					using (StreamWriter copyRow = File.AppendText(destinationFile))
					{
						copyRow.WriteLine(Row.Text);
					}
				}

			}

			//If something goes wrong catch exception and continue
			catch (Exception)
			{


			}

			//closes webpage and quits out of it everytime even if exception is caught
			driver.Close();
			driver.Quit();

		}


	}

}

