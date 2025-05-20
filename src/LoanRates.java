import java.time.Duration;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementNotInteractableException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import org.openqa.selenium.support.ui.ExpectedConditions;

import org.openqa.selenium.support.ui.WebDriverWait;

import com.poiji.bind.Poiji;

import java.io.*;

import java.util.List;

import org.openqa.selenium.support.ui.Select;

import org.apache.poi.ss.usermodel.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LoanRates {

	@SuppressWarnings("deprecation")
	public static void main(String[] args) throws InterruptedException{
		// TODO Auto-generated method stub
		List<CalculateModel> dataList = Poiji.fromExcel(new File("C:\\Users\\ninad.kulkarni\\Downloads\\SampleCalculator861.xlsx"), CalculateModel.class);

        WebDriver driver = new ChromeDriver();

//	driver.manage().window().maximize();

        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(60));

        driver.get("https://ldf--qa.sandbox.my.site.com/broker/s/calculatefinance-lease");

        Thread.sleep(5000);

        //Login

        driver.findElement(By.xpath("//div[@class='slds-form-element__control slds-grow']//input[@name='username']")).sendKeys("samiksha.bhatnagar@pivot.qa");

        driver.findElement(By.xpath("//div[@class='slds-form-element__control slds-grow']//input[@name='password']")).sendKeys("Samiksha@2");

        driver.findElement(By.xpath("//div[@class='slds-m-bottom_small slds-col slds-size_12-of-12']//div[@class='slds-align_absolute-center']//button[@type='button']")).click();
        
        System.out.println("Loan Amount\tAdvance Payments\tTerm\tRisk Band\tExpected Commission\tReceived Commission\tExpected Net Yield\tReceived Net Yield\tTotal Term in Months");
        
        for(CalculateModel data : dataList) {

            //Navigation to Home Page

            driver.findElement(By.xpath("//nav[contains(@class,'websterHomeHeader')]//a[normalize-space(text())='Home' and contains(@class,'comm-navigation__home-link')]")).click();

            //Navigating to Calculation page

            try{

                driver.findElement(By.xpath("//div[@class='themeNav']//ul[@role='menubar']//li[@role='presentation']//a[@id='3']")).click();

            }catch(ElementNotInteractableException e){

                driver.findElement(By.xpath("//div[@class='mainNavItem hasSubNav themeNavMoreLink uiMenu']//button[@class='comm-navigation__top-level-item-link comm-navigation__sub-menu-trigger js-top-level-menu-item linkBtn']")).click();

                driver.findElement(By.xpath("//a[@title='Calculate Payments / Commission']")).click();

            }

            JavascriptExecutor js = (JavascriptExecutor) driver;

            //For Loan --> Next

            WebElement element = driver.findElement(By.xpath("//div[@class='option-cont']//label[span[text()='Loan']]"));

            js.executeScript("arguments[0].click();", element);

            WebElement nextBtn = driver.findElement(By.xpath("//button[text()= 'Next']"));

            js.executeScript("arguments[0].click()", nextBtn);

            //Inputs

            //Loan Amount(Excel Sheet)
            String LoanAmount = data.getLoanAmount();
            driver.findElement(By.xpath("//div[@class='slds-form-element__control slds-grow']//input[@name='Total Loan Amount']")).sendKeys(LoanAmount);

            //Frequency (Monthly)

            driver.findElement(By.xpath("//div[@class='slds-form-element__control slds-grow']//div[@class='slds-select_container']//select[@class='slds-select']")).click();

            driver.findElement(By.xpath("//option[@value='Monthly']")).click();

            //Advance Payments(Excel Sheet)

            String advancePayments = data.getAdvancePayments();

            WebElement advancePaymentsDropdown = driver.findElement(By.xpath("//label[.//span[text()='Advance Payments']]/following-sibling::div//select"));

            Select advanceSelect = new Select(advancePaymentsDropdown);

            advanceSelect.selectByValue(advancePayments);

            //Term

            String term = data.getTerm();

            WebElement termDropdown = driver.findElement(By.xpath("//label[.//span[text()='Term']]/following-sibling::div//select"));

            Select termSelect = new Select(termDropdown);

            termSelect.selectByValue(term);

            // Select Risk Band

            String riskBand = data.getRiskBand();

            WebElement riskBandRadio = driver.findElement(By.xpath("//label[@class='slds-radio_button__label'][span[text()='" + riskBand +"']]"));

            riskBandRadio.click();

            // Set Commission

            driver.findElement(By.xpath("//label[@class='slds-radio_button__label'][span[text()='%']]")).click();

            WebElement commissionInput = driver.findElement(By.xpath("//input[@name='CommissionPer' and @inputmode='decimal']"));

            js.executeScript("arguments[0].value='" + data.getCommission() + "';", commissionInput);

            driver.findElement(By.className("calculateButton")).click();

            //Compare the inputs

            String expectedPercentage = data.getExpectedCommission();

            wait.until(ExpectedConditions.textToBePresentInElementLocated(By.xpath("//div[@class='flexForIncome']"), expectedPercentage));

            WebElement brokerPercentage = driver.findElement(By.xpath("//div[@class='flexForIncome'] "));

            String brokerTotalIncomePercentage = brokerPercentage.getText();

            String receivedCommission = brokerTotalIncomePercentage.replace("Broker Total Income", "").trim();

            data.setReceivedCommission(receivedCommission);

            //Net Yield Percentage
            String expectednetYieldPercentage = data.getExpectedNetYield();
            WebElement netYieldPercentage = driver.findElement(By.xpath("//input[@name='Net Yield%'] "));
//            wait.until(ExpectedConditions.attributeContains(netYieldPercentage, "value", expectednetYieldPercentage));
            String receivedNetYield = netYieldPercentage.getAttribute("value");
            data.setReceivedNetYield(receivedNetYield);
            if(receivedNetYield == null || receivedNetYield.isEmpty()) receivedNetYield = "null";
            
            // Total Term in Months
            WebElement totalTerm = driver.findElement(By.xpath("//input[@name='totalTermMonth']"));
            String totalTerminMonth = totalTerm.getAttribute("value");
            if(totalTerminMonth == null || totalTerminMonth.isEmpty()) totalTerminMonth = "null";
            
            System.out.println(LoanAmount+"\t\t\t"+advancePayments+"\t\t"+term+"\t"+riskBand+"\t\t\t"+expectedPercentage+"\t\t\t"+receivedCommission+"\t\t\t"+expectednetYieldPercentage+"\t\t\t"+receivedNetYield+"\t\t\t"+totalTerminMonth);
        }


        try {

            FileInputStream fis = new FileInputStream("C:\\Users\\ninad.kulkarni\\Downloads\\SampleCalculator861.xlsx");

            Workbook workbook = new XSSFWorkbook(fis);

            Sheet sheet = workbook.getSheetAt(0);

            // Find the column indexes for "Received Commission" and "Received Net Yield"

            Row headerRow = sheet.getRow(0);

            int receivedCommissionCol = -1, receivedNetYieldCol = -1;

            for (Cell cell : headerRow) {

                if (cell.getStringCellValue().equalsIgnoreCase("Received Commision")) {

                    receivedCommissionCol = cell.getColumnIndex();

                }

                if (cell.getStringCellValue().equalsIgnoreCase("Received Net Yield")) {

                    receivedNetYieldCol = cell.getColumnIndex();

                }

            }

            // Write the values row by row

            int rowIndex = 1;

            for (CalculateModel data : dataList) {

                Row row = sheet.getRow(rowIndex);

                if (row == null) row = sheet.createRow(rowIndex);

                if (receivedCommissionCol != -1) {

                    Cell cell = row.getCell(receivedCommissionCol);

                    if (cell == null) cell = row.createCell(receivedCommissionCol);

                    cell.setCellValue(data.getReceivedCommission());

                }

                if (receivedNetYieldCol != -1) {

                    Cell cell = row.getCell(receivedNetYieldCol);

                    if (cell == null) cell = row.createCell(receivedNetYieldCol);

                    cell.setCellValue(data.getReceivedNetYield());

                }

                rowIndex++;

            }

            fis.close();

            // Write the output to the file

            FileOutputStream fos = new FileOutputStream("C:\\Users\\ninad.kulkarni\\Downloads\\SampleCalculator861.xlsx");

            workbook.write(fos);

            fos.close();

            workbook.close();

        }catch(IOException e) {

            e.printStackTrace();

        }


	}

}
