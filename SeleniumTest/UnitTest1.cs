using NPOI.HSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using NPOI.SS.UserModel;
using Org.BouncyCastle.Bcpg.Attr;

namespace SeleniumTest;

[TestFixture]
public class Tests
{
    private IWebDriver driver ;
    Actions actions ;
    private static string filepath = "C:\\Automation\\CsharpAuto\\testData.xls";
    [SetUp]
    public void Setup()
    {
         driver = new ChromeDriver();
         actions = new Actions(driver);

    }

    [Test, Order(4),Category("A"),Category("B")]
    public void Test1()
    {
 
        driver.Navigate().GoToUrl("https://seleniumbase.io/demo_page");
        driver.Manage().Window.Maximize();
        IWebElement myTextInput = driver.FindElement(By.Id("myTextInput"));
        myTextInput.SendKeys("Hello World!"); 
        //actions.SendKeys(Keys.Tab).SendKeys(Keys.Tab).SendKeys(Keys.Tab).Build().Perform();
        IWebElement myButton = driver.FindElement(By.XPath("//*[@id='myButton']"));
        myButton.Click();
   
        
        actions.SendKeys(Keys.Return).Perform();
        actions.SendKeys(Keys.Return).Perform();
        
        Assert.Pass();
    }

    [Test, Order(3)]
    public void Test2()
    {
        Dictionary<string, string> tableData = new Dictionary<string, string>();
        driver.Navigate().GoToUrl("https://cosmocode.io/automation-practice-webtable");
        driver.Manage().Window.Maximize();
        IWebElement table = driver.FindElement(By.Id("countries"));
        IList<IWebElement> rows = table.FindElements(By.TagName("tr"));
        foreach (IWebElement rowElement in rows.Skip(1))
        {
            IList<IWebElement> cells = rowElement.FindElements(By.TagName("td"));
            if (cells.Count >= 2)
            {
                string countryInTable = cells[1].Text;
                string capitalInTable = cells[2].Text;
                if (tableData.Keys.Contains(countryInTable))
                {
                    countryInTable = countryInTable + "1";
                }
                tableData.Add(countryInTable, capitalInTable);

            }
        }

        Dictionary<string, string> excelData = new Dictionary<string, string>();
        using (FileStream fileStream = new FileStream(filepath, FileMode.Open, FileAccess.Read))
        {
            IWorkbook workbook= new HSSFWorkbook(fileStream);
            ISheet sheet = workbook.GetSheetAt(0);
            for (int i = 1; i <=sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    ICell cell1 = row.GetCell(0);
                    if (cell1 != null)
                    {
                        string cell1Value = cell1.StringCellValue;
                        excelData.Add(cell1Value,"");
                    }
                }
            }
        }

        using (FileStream fileStream = new FileStream(filepath, FileMode.Open, FileAccess.ReadWrite))
        {
            IWorkbook workbook = new HSSFWorkbook(fileStream);
            ISheet sheet = workbook.GetSheetAt(0);
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    ICell cell1 = row.GetCell(0);
                    ICell cell2 = row.GetCell(1)?? row.CreateCell(1);
                    if (cell1 != null)
                    {
                        string cell1Value = cell1.StringCellValue;
                        string cell2Value = tableData.ContainsKey(cell1Value) ? tableData[cell1Value] : "";
                        cell2.SetCellValue(cell2Value);
                    }
                }
            }

            using (FileStream fileStream2 = new FileStream(filepath, FileMode.Create, FileAccess.ReadWrite))
            {
                workbook.Write(fileStream2);
            }
        }
        
        Assert.Pass();
        
    }
    
    // [Test, Order(2),TestCase(@"C:\\Automation\\CsharpAuto\\scr.png"),Category("B")]
    // public void Test3(string filePath)
    // {
    //     string directory = Path.GetDirectoryName(filePath);
    //     if (!Directory.Exists(directory))
    //     {
    //         Directory.CreateDirectory(directory);
    //     }
    //
    //     Thread.Sleep(2000);
    //     Screenshot screenshot = ((ITakesScreenshot)driver).GetScreenshot();
    //     screenshot.SaveAsFile(filePath);
    //     Thread.Sleep(3000);
    //     Assert.Pass();
    // }

    [Test, Order(6)]
    public void  printParagraph()
    {
        driver.Navigate().GoToUrl("https://seleniumbase.io");
        List<IWebElement> paragraphs = new List<IWebElement>(driver.FindElements(By.TagName("p")));
        foreach (var para in paragraphs)
        {
            Console.WriteLine(para.Text);
        }
    }
    [Test, Order(1),Category("A")]
    public void  imgCount()
    {
        driver.Navigate().GoToUrl("https://seleniumbase.io");
        driver.Manage().Timeouts().ImplicitWait= TimeSpan.FromSeconds(5);
        List<IWebElement> imges = new List<IWebElement>(driver.FindElements(By.TagName("img")));
        Console.WriteLine(imges.Count);
    }

    [TearDown]
    public void CleanUp()
    {
        try
        {
            if (driver != null)
            {
                driver.Quit();
                driver.Dispose();
                driver = null;
            }
        }
        catch(Exception e)
        {
            Console.WriteLine(e);
        }
    }
}