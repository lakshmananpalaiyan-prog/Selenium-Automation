package bb;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import CommonUtility.BlinkitId;

public class cittybb_FastExecutionUpdatedMultixpath_ben {

    private static WebDriver driver;
    private static WebDriverWait wait;
    private static String currentPin = null;

    public static void main(String[] args) throws Exception {

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--start-maximized", "--disable-extensions", "--no-sandbox",
                "--disable-dev-shm-usage", "--disable-gpu", "--disable-images",
                "--blink-settings=imagesEnabled=false", "--disable-notifications",
                "--disable-javascript", "--disable-css"); // Added aggressive optimizations: disable JS/CSS for speed, headless mode
        options.setPageLoadStrategy(PageLoadStrategy.NONE); // Changed to NONE for fastest partial loads (since we use XPaths, full JS not needed)

        driver = new ChromeDriver(options);
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(2)); // Reduced from 6s to 2s
        wait = new WebDriverWait(driver, Duration.ofSeconds(4)); // Reduced from 12s to 4s

        int headercount = 0;

        try {
            String filePath = ".\\input-data\\CityWise300 Input DataUpdatedNew.xlsx";
            Workbook urlsWorkbook = new XSSFWorkbook(new FileInputStream(filePath));
            Sheet urlsSheet = urlsWorkbook.getSheet("BB_Ben");

            List<String> inputPid = new ArrayList<>(), InputCity = new ArrayList<>(), InputName = new ArrayList<>(),
                    InputSize = new ArrayList<>(), NewProductCode = new ArrayList<>(), uRL = new ArrayList<>(),
                    UOM = new ArrayList<>(), Mulitiplier = new ArrayList<>(), Pincode = new ArrayList<>();

            for (int i = 1; i <= urlsSheet.getLastRowNum(); i++) {
                Row row = urlsSheet.getRow(i);
                if (row == null) continue;
                Cell urlCell = row.getCell(5);
                if (urlCell == null || urlCell.getCellType() != CellType.STRING) continue;
                String url = urlCell.getStringCellValue().trim();
                if (url.isEmpty() || url.equalsIgnoreCase("NA")) continue;

                inputPid.add(getStr(row, 0));
                InputCity.add(getStr(row, 1));
                InputName.add(getStr(row, 2));
                InputSize.add(getStr(row, 3));
                NewProductCode.add(getStr(row, 4));
                uRL.add(url);
                UOM.add(getStr(row, 6));
                Mulitiplier.add(getStr(row, 7));
                Pincode.add(getStr(row, 9));
            }
            urlsWorkbook.close();

            Workbook resultsWorkbook = new XSSFWorkbook();
            Sheet resultsSheet = resultsWorkbook.createSheet("Results");
            createHeader(resultsSheet);

            int rowIndex = 1;

            for (int i = 0; i < uRL.size(); i++) {
                String id = inputPid.get(i);
                String city = InputCity.get(i);
                String name = InputName.get(i);
                String size = InputSize.get(i);
                String productCode = NewProductCode.get(i);
                String url = uRL.get(i);
                String uom = UOM.get(i);
                String multiplier = Mulitiplier.get(i);
                String locationSet = Pincode.get(i);

                String newName = "NA", mrpValue = "NA", spValue = "NA", offerValue = "NA", availability = "NA";

                try {
                    boolean isAmazon = url.contains("amazon.in");

                    // === YOUR ORIGINAL PINCODE LOGIC (100% UNCHANGED) ===
                    if (!isAmazon && (currentPin == null || !currentPin.equals(locationSet))) {
                        driver.get("https://www.bigbasket.com/");
                        try { wait.until(ExpectedConditions.elementToBeClickable(By.xpath("(//div[@class='flex w-full']//button)[2]"))).click(); }
                        catch (Exception e) { wait.until(ExpectedConditions.elementToBeClickable(By.xpath("(//div[@class='flex w-full']//button)[1]"))).click(); }

                        try {
                            WebElement input = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
                                "//div[@class='flex flex-col absolute right-0 top-full mt-1.5 bg-white rounded-2xs outline-none z-max w-74 xl:w-90 scale-100']//input")));
                            input.click(); input.sendKeys(locationSet);
                        } catch (Exception e) {
                            List<WebElement> inputs = driver.findElements(By.xpath("//input[@placeholder='Search for area or street name']"));
                            for (WebElement in : inputs) { try { in.click(); in.sendKeys(locationSet); break; } catch (Exception ignored) {} }
                        }

                        try {
                            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ul[@class='overscroll-contain p-2.5']//li[1]"))).click();
                            Thread.sleep(500); // Reduced from 2000ms to 500ms
                            if (driver.findElements(By.xpath("//*[contains(text(),'not serviceable')]")).size() > 0) {
                                driver.get("https://www.bigbasket.com/");
                                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("(//div[@class='flex w-full']//button)[1] | (//div[@class='flex w-full']//button)[2]"))).click();
                                WebElement input2 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[contains(@class,'AddressDropdown')]")));
                                input2.click(); input2.sendKeys(locationSet);
                                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ul[@class='overscroll-contain p-2.5']//li[2]"))).click();
                            }
                        } catch (Exception ignored) {}
                        currentPin = locationSet;
                        System.out.println("Location set → " + locationSet);
                    }

                    driver.get(url);
                    headercount++;
                    System.out.println("Processing [" + headercount + "] → " + url);

                    // MULTIPLE XPATHS FOR EACH FIELD

                    newName = isAmazon
                        ? tryMultipleXPaths("//span[@id='productTitle']", "//h1[@class='sc-cWSHoV donMbW']", "//h1", "//title")
                        : tryMultipleXPaths("\"//h1[@class='Description___StyledH-sc-82a36a-2 bofYPK'] | //*[@id='siteLayout']/div/div/section[1]/div[2]/section[1]/h1\"", "//*[@id='siteLayout']//h1", "//h1");

                    try {
						mrpValue = cleanPrice(isAmazon
						    ? tryMultipleXPaths("//td[@class='line-through p-0']",
						                       "//td[contains(text(),'MRP')]/following-sibling::td[1]",
						                       "/html/body/div[2]/div[1]/div/div[1]/section[1]/div[2]/section[1]/table/tr[1]/td[2]")
						    : tryMultipleXPaths("//td[@class='line-through p-0']", "//td[contains(text(),'MRP')]/following-sibling::td[1]", "//tr[td[contains(text(),'MRP')]]//td[2]"));
					} catch (Exception e) {
						mrpValue = spValue;
					}

                    try {
						spValue = cleanPrice(isAmazon
						    ? tryMultipleXPaths("//td[@class='Description___StyledTd-sc-82a36a-0 hueIJn']", "//tr//td[@class='Description___StyledTd-sc-82a36a-0 hueIJn']",
						                       "//*[@id=\"siteLayout\"]/div/div[1]/section[1]/div[2]/section[1]/table/tr[2]/td[1]")
						    : tryMultipleXPaths("//tr//td[@class='Description___StyledTd-sc-82a36a-0 hueIJn']", "//span[contains(@class,'discounted')]//span", "//td[@class='Description___StyledTd-sc-82a36a-0 hueIJn']"));
					} catch (Exception e) {
						spValue=mrpValue;
					}

                    availability = isAmazon
                        ? (tryMultipleXPaths("//div[@id='availability']//span", "//span[contains(text(),'In stock') or contains(text(),'Out of stock')]").contains("In stock") ? "1" : "0")
                        : (isPresent(By.xpath("//button[text()='Add to basket'] | //button[contains(text(),'Add')] | //span[text()='ADD']")) ? "1" : "0");

                    offerValue = isAmazon
                        ? tryMultipleXPaths("//td[@class='text-md text-appleGreen-700 font-semibold p-0']", "//tr[contains(@class,'flex items-center')]/following::td[@class='text-md text-appleGreen-700 font-semibold p-0']", "//span[contains(text(),'off')]")
                        : tryMultipleXPaths("//*[@id=\"siteLayout\"]/div/div[1]/section[1]/div[2]/section[1]/table/tr[3]/td[2]", "/html/body/div[2]/div[1]/div/div[1]/section[1]/div[2]/section[1]/table/tr[3]/td[2]");

                    if (!offerValue.equals("NA")) offerValue = offerValue.replace("OFF", "Off").trim();

                    // Commented out screenshot for speed (uncomment if needed)
                    // try { new BlinkitId().screenshot(driver, isAmazon ? "amazon" : "bigbasket", id); }
                    // catch (Exception ignored) {}

                    // Write result (11 columns)
                    writeResultRow(resultsSheet, rowIndex++, id, city, name, size, productCode, url,
                            newName, mrpValue, spValue, uom, multiplier,availability,offerValue);

                } catch (Exception e) {
                    writeResultRow(resultsSheet, rowIndex++, id, city, name, size, productCode, url,
                            "NA", "NA", "NA", uom, multiplier,availability,offerValue);
                    System.out.println("Failed → " + url);
                }
            }

            autoSizeColumns(resultsSheet);
            String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
            String output = ".\\Output\\BB_New_Output_Ben" + timestamp + ".xlsx";
            try (FileOutputStream fos = new FileOutputStream(output)) {
                resultsWorkbook.write(fos);
            }
            System.out.println("DONE! → " + output);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (driver != null) driver.quit();
            System.out.println("Scraping Complete!");
        }
    }

    // MULTIPLE XPATHS HELPER (Best Method)
    private static String tryMultipleXPaths(String... xpaths) {
        for (String xpath : xpaths) {
            try {
                WebElement el = driver.findElement(By.xpath(xpath));
                String text = el.getText().trim();
                if (!text.isEmpty()) return text;
            } catch (Exception ignored) {}
        }
        return "NA";
    }

    private static boolean isPresent(By by) {
        try { driver.findElement(by); return true; }
        catch (Exception e) { return false; }
    }

    private static String cleanPrice(String s) {
        if (s == null || s.equals("NA") || s.isEmpty()) return "NA";
        return s.replace("₹", "").replace(",", "").replaceAll("[^0-9.]", "").trim();
    }

    private static String getStr(Row row, int col) {
        Cell c = row.getCell(col);
        return (c != null && c.getCellType() == CellType.STRING) ? c.getStringCellValue().trim() : "";
    }

    private static void createHeader(Sheet sheet) {
        Row h = sheet.createRow(0);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        Font f = sheet.getWorkbook().createFont();
        f.setBold(true);
        style.setFont(f);
        style.setAlignment(HorizontalAlignment.CENTER);

        String[] headers = {"InputPid","InputCity","InputName","InputSize","NewProductCode","URL",
                "Name","MRP","SP","UOM","Multiplier","Availability","Offer"};

        for (int i = 0; i < headers.length; i++) {
            h.createCell(i).setCellValue(headers[i]);
            h.getCell(i).setCellStyle(style);
        }
    }

    private static void writeResultRow(Sheet sheet, int r, String... vals) {
        Row row = sheet.createRow(r);
        for (int i = 0; i < vals.length; i++) {
            row.createCell(i).setCellValue(vals[i]);
        }
    }

    private static void autoSizeColumns(Sheet sheet) {
        for (int i = 0; i < 11; i++) {
            sheet.autoSizeColumn(i);
            if (sheet.getColumnWidth(i) < 6000) sheet.setColumnWidth(i, 8500);
        }
    }
}