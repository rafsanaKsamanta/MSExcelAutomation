import io.appium.java_client.windows.WindowsDriver;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.net.URL;
import java.time.LocalDate;
import java.util.concurrent.TimeUnit;

// class create and give appropriate name

public class MSExcel {
    private static WindowsDriver excel = null;

    public static String getDate() {
        LocalDate date = LocalDate.now();
        return date.toString();
    }

    @BeforeClass
    public static void setUp() {
        try {
            DesiredCapabilities capabilities = new DesiredCapabilities();
            capabilities.setCapability("app", "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL"); // define path of the application
            capabilities.setCapability("platformName", "Windows");
            capabilities.setCapability("deviceName", "WindowsPC");
            excel = new WindowsDriver(new URL("http://127.0.0.1:4723"), capabilities); // appium server
            excel.manage().timeouts().implicitlyWait(2,TimeUnit.SECONDS); // sleep time
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    //write test cases from here

    @Test(priority = '1', description = "Open a blank page from excel", testName = "TC_Ex_01")
    public void openBlankPage() {


//        WebElement blankPage = excel.findElementByName("Blank workbook");
        Actions actions = new Actions(excel);
        actions.keyDown(Keys.ALT)
                .sendKeys("f")
                .sendKeys("l")
                .keyUp(Keys.ALT)
                .perform();
    }

    @Test(priority = '2',description = "Get value for column1",testName = "TC_Ex_02")
    public void getColumnValue(){
        excel.findElementByName("A1").click();
        excel.findElementByName("A1").sendKeys("25");
        WebElement pressEnter1=excel.findElementByName("B1");
        pressEnter1.sendKeys(Keys.ARROW_RIGHT);
        excel.findElementByName("B1").sendKeys("55");


    }
    @Test(priority = '3',description = "Get value for column2",testName = "TC_Ex_03")
    public void SUM(){

        WebElement pressEnter1=excel.findElementByName("B1");
        pressEnter1.sendKeys(Keys.ARROW_RIGHT);

        excel.findElementByName("Sum").click();

        Actions actions = new Actions(excel);
        actions.keyDown(Keys.ALT)
                .sendKeys("=")
                .keyUp(Keys.ALT)
                .perform();
    }

    @Test(priority = '4',description = "File save",testName = "TC_Ex_04")
    public void fileSave(){

        WebElement save1=excel.findElementByName("File Tab");
        save1.click();

        WebElement save2 = excel.findElementByName("Save");
        save2.click();

        WebElement desktop = excel.findElementByName("Desktop");
        desktop.click();

        excel.findElementByName("Save").click();

        WebElement pressEnter1=excel.findElementByName("Save");
        pressEnter1.sendKeys(Keys.ENTER);


    }



}
