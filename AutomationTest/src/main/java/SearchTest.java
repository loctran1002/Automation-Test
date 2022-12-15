import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.interactions.Actions;
import org.testng.Assert;
import org.testng.annotations.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

public class SearchTest
{
    private final Path path = Path.of("src/test/resources/TestSuite.xlsx");
    private FileInputStream fin;
    private XSSFWorkbook workbook;
    private XSSFSheet sheetSummary;
    private XSSFSheet sheetInsert;
    private XSSFSheet sheetSearch;
    private File file;
    private WebDriver webDriver;
    private static int index;
    private static int column;
    private static boolean status;

    @BeforeSuite
    public void ClearOldData() throws IOException
    {
        file = new File(path.toAbsolutePath().toString());
        fin = new FileInputStream(file);
        workbook = new XSSFWorkbook(fin);
        sheetSummary = workbook.getSheetAt(2);
        sheetInsert = workbook.getSheetAt(3);
        sheetSearch = workbook.getSheetAt(4);

        // Remove old version browser
        sheetSummary.getRow(2).getCell(0).setCellValue("");
        sheetSummary.getRow(3).getCell(0).setCellValue("");
        sheetInsert.getRow(7).getCell(0).setCellValue("");
        sheetInsert.getRow(8).getCell(0).setCellValue("");
        sheetSearch.getRow(7).getCell(0).setCellValue("");
        sheetSearch.getRow(8).getCell(0).setCellValue("");

        // Setup default statistic 3 sheets
        for (int i = 1; i < 7; i++)
        {
            sheetInsert.getRow(7).getCell(i).setCellValue(0);
            sheetInsert.getRow(8).getCell(i).setCellValue(0);
            sheetSearch.getRow(7).getCell(i).setCellValue(0);
            sheetSearch.getRow(8).getCell(i).setCellValue(0);
            sheetSummary.getRow(2).getCell(i).setCellValue(0);
            sheetSummary.getRow(3).getCell(i).setCellValue(0);
        }

        // Setup default status sheet Insert
        for (int i = 18; i < 73; i += 3)
        {
            sheetInsert.getRow(i).getCell(9).setCellValue("");
            sheetInsert.getRow(i).getCell(10).setCellValue("");
        }

        // Setup default status sheet Search
        for (int i = 18; i < 55; i += 3)
        {
            sheetSearch.getRow(i).getCell(9).setCellValue("");
            sheetSearch.getRow(i).getCell(10).setCellValue("");
        }

        // Setup default status sheet Summary
        for (int i = 9; i < 42; i++)
        {
            if (i == 28)
                continue;
            sheetSummary.getRow(i).getCell(4).setCellValue("");
            sheetSummary.getRow(i).getCell(5).setCellValue("");
        }

        workbook.write(new FileOutputStream(file));
    }

    @Parameters({"browser"})
    @BeforeClass
    public void PrepareData(String browser) throws IOException
    {
        index = 0;
        file = new File(path.toAbsolutePath().toString());
        fin = new FileInputStream(file);
        workbook = new XSSFWorkbook(fin);
        sheetSummary = workbook.getSheetAt(2);
        sheetInsert = workbook.getSheetAt(3);
        sheetSearch = workbook.getSheetAt(4);

        // Write version
        if (browser.equals("chrome"))
        {
            webDriver = new ChromeDriver();
            String version = "Google Chrome ";
            version += ((ChromeDriver) webDriver).getCapabilities().getBrowserVersion().toString();
            sheetSummary.getRow(2).getCell(0).setCellValue(version);
            sheetInsert.getRow(7).getCell(0).setCellValue(version);
            sheetSearch.getRow(7).getCell(0).setCellValue(version);
        }
        else if (browser.equals("edge"))
        {
            webDriver = new EdgeDriver();
            String version = "Microsoft Edge ";
            version += ((EdgeDriver) webDriver).getCapabilities().getBrowserVersion().toString();
            sheetSummary.getRow(3).getCell(0).setCellValue(version);
            sheetInsert.getRow(8).getCell(0).setCellValue(version);
            sheetSearch.getRow(8).getCell(0).setCellValue(version);
        }

        webDriver.quit();
        workbook.write(new FileOutputStream(file));
    }

    @Parameters({"browser"})
    @AfterClass
    public void Report(String browser) throws IOException
    {
        int passed = 0, failed = 0;
        int row = browser.equals("chrome") ? 7 : 8;
        column = browser.equals("chrome") ? 9 : 10;
        // Sheet Insert
        for (int i = 18; i < 55; i += 3)
        {
            if (sheetSearch.getRow(i).getCell(column).toString().equals("Passed"))
                passed++;
            else if (sheetSearch.getRow(i).getCell(column).toString().equals("Failed"))
                failed++;
        }
        for (int i = 0; i < 6; i++)
        {
            String test = sheetSearch.getRow(6).getCell(i + 1).toString();
            if (test.equals("Passed"))
                sheetSearch.getRow(row).getCell(i + 1).setCellValue(passed);
            else if (test.equals("Failed"))
                sheetSearch.getRow(row).getCell(i + 1).setCellValue(failed);
            else
                sheetSearch.getRow(row).getCell(i + 1).setCellValue(0);
        }

        // Sheet Summary
        row = browser.equals("chrome") ? 2 : 3;
        column = browser.equals("chrome") ? 4 : 5;
        passed = 0;
        failed = 0;
        for (int i = 9; i < 42; i++)
        {
            if (i == 28)
                continue;
            if (sheetSummary.getRow(i).getCell(column).getRawValue().equals("Passed"))
                passed++;
            else if (sheetSummary.getRow(i).getCell(column).getRawValue().equals("Failed"))
                failed++;
        }
        for (int i = 0; i < 6; i++)
        {
            String test = sheetSummary.getRow(1).getCell(i + 1).toString();
            if (test.equals("Passed"))
                sheetSummary.getRow(row).getCell(i + 1).setCellValue(passed);
            else if (test.equals("Failed"))
                sheetSummary.getRow(row).getCell(i + 1).setCellValue(failed);
            else
                sheetSummary.getRow(row).getCell(i + 1).setCellValue(0);
        }

        workbook.write(new FileOutputStream(file));
    }

    @Parameters({"browser"})
    @BeforeMethod
    public void Setup(String browser)
    {
        if (browser.equals("chrome"))
        {
            WebDriverManager.chromedriver().setup();
            webDriver = new ChromeDriver();
        }
        if (browser.equals("edge"))
        {
            WebDriverManager.edgedriver().setup();
            webDriver = new EdgeDriver();
        }
        webDriver.manage().window().maximize();
        webDriver.get("https://visualgo.net/en/hashtable");
        Sleep(100);
        webDriver.findElement(By.id("gdpr-reject")).click();
        Sleep(100);
        webDriver.findElement(By.className("electure-end")).click();
        Sleep(100);

        WebElement draggable = webDriver.findElement(By.xpath("//div[@id='speed-input']//span"));
        WebElement droppable = webDriver.findElement(By.id("viz-speed-value"));
        new Actions(webDriver)
                .dragAndDrop(draggable, droppable)
                .perform();

        webDriver.findElement(By.id("create")).click();
        Sleep(100);
        webDriver.findElement(By.id("v-createM")).clear();
        String tableSize = sheetSearch.getRow(18 + index * 3).getCell(3).getRawValue();
        webDriver.findElement(By.id("v-createM")).sendKeys(tableSize);
        Sleep(100);
        webDriver.findElement(By.id("v-createN")).clear();
        Sleep(100);
        webDriver.findElement(By.id("create-go")).click();
        Sleep(100);

        String init = sheetSearch.getRow(18 + index * 3).getCell(4).toString();
        if (!init.equals("None"))
        {
            webDriver.findElement(By.id("insert")).click();
            Sleep(100);
            webDriver.findElement(By.id("v-insert")).clear();
            webDriver.findElement(By.id("v-insert")).sendKeys(init);
            webDriver.findElement(By.id("insert-go")).click();
            Sleep(200 * init.length());
        }
        index++;
        status = false;
    }

    @Parameters({"browser"})
    @AfterMethod
    public void Cleanup(String browser) throws IOException
    {
        String result = (status) ? "Passed" : "Failed";
        column = browser.equals("chrome") ? 9 : 10;
        sheetSearch.getRow(18 + (index - 1) * 3).getCell(column).setCellValue(result);
        column = browser.equals("chrome") ? 4 : 5;
        sheetSummary.getRow(29 + (index - 1)).getCell(column).setCellValue(result);
        workbook.write(new FileOutputStream(file));
        webDriver.quit();
    }

    @Test
    public void SEA_001()
    {
        String str = sheetSearch.getRow(19 + (index - 1) * 3).getCell(6).toString();
        String input = str.substring(16);
        String checkNoInput = "kh\u00f4ng nh\u1eadp";
        webDriver.findElement(By.id("actions-hide")).click();
        Sleep(100);
        webDriver.findElement(By.id("search")).click();
        Sleep(100);
        webDriver.findElement(By.id("v-search")).clear();
        if (!str.contains(checkNoInput))
            webDriver.findElement(By.id("v-search")).sendKeys(input);

        String ExpectedTitle = sheetSearch.getRow(18 + (index - 1) * 3).getCell(7).toString();
        if (ExpectedTitle.equals("Nothing responses"))
            ExpectedTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();

        webDriver.findElement(By.id("search-go")).click();
        Sleep(2000);
        String ActualTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();;
        Assert.assertEquals(ActualTitle, ExpectedTitle);
        status = true;
    }

    @Test
    public void SEA_002()
    {
        String str = sheetSearch.getRow(19 + (index - 1) * 3).getCell(6).toString();
        String input = str.substring(16);
        String checkNoInput = "kh\u00f4ng nh\u1eadp";
        webDriver.findElement(By.id("actions-hide")).click();
        Sleep(100);
        webDriver.findElement(By.id("search")).click();
        Sleep(100);
        webDriver.findElement(By.id("v-search")).clear();
        if (!str.contains(checkNoInput))
            webDriver.findElement(By.id("v-search")).sendKeys(input);

        String ExpectedTitle = sheetSearch.getRow(18 + (index - 1) * 3).getCell(7).toString();
        if (ExpectedTitle.equals("Nothing responses"))
            ExpectedTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();

        webDriver.findElement(By.id("search-go")).click();
        Sleep(2000);
        String ActualTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();;
        Assert.assertEquals(ActualTitle, ExpectedTitle);
        status = true;
    }

    @Test
    public void SEA_003()
    {
        String str = sheetSearch.getRow(19 + (index - 1) * 3).getCell(6).toString();
        String input = str.substring(16);
        String checkNoInput = "kh\u00f4ng nh\u1eadp";
        webDriver.findElement(By.id("actions-hide")).click();
        Sleep(100);
        webDriver.findElement(By.id("search")).click();
        Sleep(100);
        webDriver.findElement(By.id("v-search")).clear();
        if (!str.contains(checkNoInput))
            webDriver.findElement(By.id("v-search")).sendKeys(input);

        String ExpectedTitle = sheetSearch.getRow(18 + (index - 1) * 3).getCell(7).toString();
        if (ExpectedTitle.equals("Nothing responses"))
            ExpectedTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();

        webDriver.findElement(By.id("search-go")).click();
        Sleep(2000);
        String ActualTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();;
        Assert.assertEquals(ActualTitle, ExpectedTitle);
        status = true;
    }

    @Test
    public void SEA_004()
    {
        String str = sheetSearch.getRow(19 + (index - 1) * 3).getCell(6).toString();
        String input = str.substring(16);
        String checkNoInput = "kh\u00f4ng nh\u1eadp";
        webDriver.findElement(By.id("actions-hide")).click();
        Sleep(100);
        webDriver.findElement(By.id("search")).click();
        Sleep(100);
        webDriver.findElement(By.id("v-search")).clear();
        if (!str.contains(checkNoInput))
            webDriver.findElement(By.id("v-search")).sendKeys(input);

        String ExpectedTitle = sheetSearch.getRow(18 + (index - 1) * 3).getCell(7).toString();
        if (ExpectedTitle.equals("Nothing responses"))
            ExpectedTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();

        webDriver.findElement(By.id("search-go")).click();
        Sleep(2000);
        String ActualTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();;
        Assert.assertEquals(ActualTitle, ExpectedTitle);
        status = true;
    }

    @Test
    public void SEA_005()
    {
        String str = sheetSearch.getRow(19 + (index - 1) * 3).getCell(6).toString();
        String input = str.substring(16);
        String checkNoInput = "kh\u00f4ng nh\u1eadp";
        webDriver.findElement(By.id("actions-hide")).click();
        Sleep(100);
        webDriver.findElement(By.id("search")).click();
        Sleep(100);
        webDriver.findElement(By.id("v-search")).clear();
        if (!str.contains(checkNoInput))
            webDriver.findElement(By.id("v-search")).sendKeys(input);

        String ExpectedTitle = sheetSearch.getRow(18 + (index - 1) * 3).getCell(7).toString();
        if (ExpectedTitle.equals("Nothing responses"))
            ExpectedTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();

        webDriver.findElement(By.id("search-go")).click();
        Sleep(2000);
        String ActualTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();;
        Assert.assertEquals(ActualTitle, ExpectedTitle);
        status = true;
    }

    @Test
    public void SEA_006()
    {
        String str = sheetSearch.getRow(19 + (index - 1) * 3).getCell(6).toString();
        String input = str.substring(16);
        String checkNoInput = "kh\u00f4ng nh\u1eadp";
        webDriver.findElement(By.id("actions-hide")).click();
        Sleep(100);
        webDriver.findElement(By.id("search")).click();
        Sleep(100);
        webDriver.findElement(By.id("v-search")).clear();
        if (!str.contains(checkNoInput))
            webDriver.findElement(By.id("v-search")).sendKeys(input);

        String ExpectedTitle = sheetSearch.getRow(18 + (index - 1) * 3).getCell(7).toString();
        if (ExpectedTitle.equals("Nothing responses"))
            ExpectedTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();

        webDriver.findElement(By.id("search-go")).click();
        Sleep(2000);
        String ActualTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();;
        Assert.assertEquals(ActualTitle, ExpectedTitle);
        status = true;
    }

    @Test
    public void SEA_007()
    {
        String str = sheetSearch.getRow(19 + (index - 1) * 3).getCell(6).toString();
        String input = str.substring(16);
        String checkNoInput = "kh\u00f4ng nh\u1eadp";
        webDriver.findElement(By.id("actions-hide")).click();
        Sleep(100);
        webDriver.findElement(By.id("search")).click();
        Sleep(100);
        webDriver.findElement(By.id("v-search")).clear();
        if (!str.contains(checkNoInput))
            webDriver.findElement(By.id("v-search")).sendKeys(input);

        String ExpectedTitle = sheetSearch.getRow(18 + (index - 1) * 3).getCell(7).toString();
        if (ExpectedTitle.equals("Nothing responses"))
            ExpectedTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();

        webDriver.findElement(By.id("search-go")).click();
        Sleep(2000);
        String ActualTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();;
        Assert.assertEquals(ActualTitle, ExpectedTitle);
        status = true;
    }

    @Test
    public void SEA_008()
    {
        String str = sheetSearch.getRow(19 + (index - 1) * 3).getCell(6).toString();
        String input = str.substring(16);
        String checkNoInput = "kh\u00f4ng nh\u1eadp";
        webDriver.findElement(By.id("actions-hide")).click();
        Sleep(100);
        webDriver.findElement(By.id("search")).click();
        Sleep(100);
        webDriver.findElement(By.id("v-search")).clear();
        if (!str.contains(checkNoInput))
            webDriver.findElement(By.id("v-search")).sendKeys(input);

        String ExpectedTitle = sheetSearch.getRow(18 + (index - 1) * 3).getCell(7).toString();
        if (ExpectedTitle.equals("Nothing responses"))
            ExpectedTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();

        webDriver.findElement(By.id("search-go")).click();
        Sleep(2000);
        String ActualTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();;
        Assert.assertEquals(ActualTitle, ExpectedTitle);
        status = true;
    }

    @Test
    public void SEA_009()
    {
        String str = sheetSearch.getRow(19 + (index - 1) * 3).getCell(6).toString();
        String input = str.substring(16);
        String checkNoInput = "kh\u00f4ng nh\u1eadp";
        webDriver.findElement(By.id("actions-hide")).click();
        Sleep(100);
        webDriver.findElement(By.id("search")).click();
        Sleep(100);
        webDriver.findElement(By.id("v-search")).clear();
        if (!str.contains(checkNoInput))
            webDriver.findElement(By.id("v-search")).sendKeys(input);

        String ExpectedTitle = sheetSearch.getRow(18 + (index - 1) * 3).getCell(7).toString();
        if (ExpectedTitle.equals("Nothing responses"))
            ExpectedTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();

        webDriver.findElement(By.id("search-go")).click();
        Sleep(2000);
        String ActualTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();;
        Assert.assertEquals(ActualTitle, ExpectedTitle);
        status = true;
    }

    @Test
    public void SEA_010()
    {
        String str = sheetSearch.getRow(19 + (index - 1) * 3).getCell(6).toString();
        String input = str.substring(16);
        String checkNoInput = "kh\u00f4ng nh\u1eadp";
        webDriver.findElement(By.id("actions-hide")).click();
        Sleep(100);
        webDriver.findElement(By.id("search")).click();
        Sleep(100);
        webDriver.findElement(By.id("v-search")).clear();
        if (!str.contains(checkNoInput))
            webDriver.findElement(By.id("v-search")).sendKeys(input);

        String ExpectedTitle = sheetSearch.getRow(18 + (index - 1) * 3).getCell(7).toString();
        if (ExpectedTitle.equals("Nothing responses"))
            ExpectedTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();

        webDriver.findElement(By.id("search-go")).click();
        Sleep(2000);
        String ActualTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();;
        Assert.assertEquals(ActualTitle, ExpectedTitle);
        status = true;
    }

    @Test
    public void SEA_011()
    {
        String str = sheetSearch.getRow(19 + (index - 1) * 3).getCell(6).toString();
        String input = str.substring(16);
        String checkNoInput = "kh\u00f4ng nh\u1eadp";
        webDriver.findElement(By.id("actions-hide")).click();
        Sleep(100);
        webDriver.findElement(By.id("search")).click();
        Sleep(100);
        webDriver.findElement(By.id("v-search")).clear();
        if (!str.contains(checkNoInput))
            webDriver.findElement(By.id("v-search")).sendKeys(input);

        String ExpectedTitle = sheetSearch.getRow(18 + (index - 1) * 3).getCell(7).toString();
        if (ExpectedTitle.equals("Nothing responses"))
            ExpectedTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();

        webDriver.findElement(By.id("search-go")).click();
        Sleep(2000);
        String ActualTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();;
        Assert.assertEquals(ActualTitle, ExpectedTitle);
        status = true;
    }

    @Test
    public void SEA_012()
    {
        String str = sheetSearch.getRow(19 + (index - 1) * 3).getCell(6).toString();
        String input = str.substring(16);
        String checkNoInput = "kh\u00f4ng nh\u1eadp";
        webDriver.findElement(By.id("actions-hide")).click();
        Sleep(100);
        webDriver.findElement(By.id("search")).click();
        Sleep(100);
        webDriver.findElement(By.id("v-search")).clear();
        if (!str.contains(checkNoInput))
            webDriver.findElement(By.id("v-search")).sendKeys(input);

        String ExpectedTitle = sheetSearch.getRow(18 + (index - 1) * 3).getCell(7).toString();
        if (ExpectedTitle.equals("Nothing responses"))
            ExpectedTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();

        webDriver.findElement(By.id("search-go")).click();
        Sleep(2000);
        String ActualTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();;
        Assert.assertEquals(ActualTitle, ExpectedTitle);
        status = true;
    }

    @Test
    public void SEA_013()
    {
        String str = sheetSearch.getRow(19 + (index - 1) * 3).getCell(6).toString();
        String input = str.substring(16);
        String checkNoInput = "kh\u00f4ng nh\u1eadp";
        webDriver.findElement(By.id("actions-hide")).click();
        Sleep(100);
        webDriver.findElement(By.id("search")).click();
        Sleep(100);
        webDriver.findElement(By.id("v-search")).clear();
        if (!str.contains(checkNoInput))
            webDriver.findElement(By.id("v-search")).sendKeys(input);

        String ExpectedTitle = sheetSearch.getRow(18 + (index - 1) * 3).getCell(7).toString();
        if (ExpectedTitle.equals("Nothing responses"))
            ExpectedTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();

        webDriver.findElement(By.id("search-go")).click();
        Sleep(2000);
        String ActualTitle = webDriver.findElement(By.xpath("//div[@id='status']//p")).getText();;
        Assert.assertEquals(ActualTitle, ExpectedTitle);
        status = true;
    }

    private void Sleep(int time)
    {
        try
        {
            Thread.sleep(time);
        }
        catch (Exception e)
        {
            System.getLogger(e.toString());
        }
    }
}
