import com.codeborne.selenide.ElementsCollection;
import com.codeborne.selenide.Screenshots;
import com.codeborne.selenide.WebDriverRunner;
import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.http.util.Asserts;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.Capabilities;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

import static com.codeborne.selenide.Selenide.*;
import static org.junit.jupiter.api.Assertions.assertEquals;


public class YandexTest {

    @Test
    public void YandexTest() throws IOException {
        //Считываем значения
        FileInputStream inputStream = new FileInputStream(new File("TestData.xls"));
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        HSSFSheet sheet = workbook.getSheetAt(0);
        String yandexAdress = (sheet.getRow(1).getCell(1)).toString();
        String cBRFSerchingString = (sheet.getRow(2).getCell(1)).toString();
        String firstResultFoundXpath = (sheet.getRow(3).getCell(1)).toString();
        String CBRFTitle = (sheet.getRow(4).getCell(1)).toString();
        String keyRateXpath = (sheet.getRow(5).getCell(1)).toString();
        String searchingFieldXpath = (sheet.getRow(6).getCell(1)).toString();
        String headlinesXpath = (sheet.getRow(7).getCell(1)).toString();
        String webAddressesXpath = (sheet.getRow(8).getCell(1)).toString();
        workbook.close();
        //Зайти на yandex.ru. В поисковую строку ввести «цб рф»
        ChromeDriver driver = new ChromeDriver();
//      InternetExplorerDriver driver = new InternetExplorerDriver();//TODO: разобраться с настройками InternetExplorerDriver
        WebDriverRunner.setWebDriver(driver);
        driver.manage().window().maximize();
        open(yandexAdress);
        $(By.xpath(searchingFieldXpath)).val(cBRFSerchingString).pressEnter();
        //Нажать на 1 результат
        $(By.xpath(firstResultFoundXpath)).click();
        //Проверить нахождение на сайте Центрального Банка.
        switchTo().window(CBRFTitle);
        assertEquals(CBRFTitle, title());
        //Закрыть вкладку с яндексом.
        switchTo().window(0).close();
        //Найти и сохранить ключевую ставку
        switchTo().window(CBRFTitle);
        String keyRate = $(By.xpath(keyRateXpath)).text();
        //Открыть новую вкладку, зайти на yandex.ru.
        actions().sendKeys(Keys.CONTROL + "t");
        open(yandexAdress);
        //В поисковую строку ввести сохраненную ключевую ставку
        $(By.xpath(searchingFieldXpath)).val(keyRate).pressEnter();
        //Сохранить заголовки и ссылки первых трёх результатов в текстовый документ. (документ назвать «новости.txt»)
        List<String> headlines = $$(By.xpath(headlinesXpath)).first(3).texts();
        List<String> webAddresses = $$(By.xpath(webAddressesXpath)).first(3).texts();
        //Запись в файл
        try {
//            String filePath = new File("").getAbsolutePath();
            FileWriter writer = new FileWriter("Новости.txt", false);
            for (int x = 0; x < 3; x = x + 1) {
                String headlineText = headlines.get(x);
                writer.write(headlineText);
                writer.append('\n');
                String adressText = webAddresses.get(x);
                writer.write(adressText);
                writer.append('\n');
                writer.append('\n');
            }
            writer.flush();
        } catch (IOException ex) {

            System.out.println(ex.getMessage());
        }
        //Сделать скриншот страницы, скриншот поместить в ту же папку, что и текстовый документ.
        screenshot("News");//TODO: скриншот поместить в ту же папку, что и текстовый документ
        //Закрыть браузер.
        driver.quit();
    }
}
