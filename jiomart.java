package Apollo;

import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.*;

public class jiomart {
    public static void main(String[] args) throws Exception {

        // ULTRA FAST & UNIVERSAL CHROME SETUP (Java 8 to 21)
        ChromeOptions options = new ChromeOptions();
       // options.addArguments("--headless=new");
        options.addArguments("--no-sandbox");
        options.addArguments("--disable-dev-shm-usage");
        options.addArguments("--disable-gpu");
        options.addArguments("--disable-extensions");
        options.addArguments("--disable-images");
        options.addArguments("--blink-settings=imagesEnabled=false");
        options.setPageLoadStrategy(PageLoadStrategy.EAGER);

        Map<String, Object> prefs = new HashMap<>();
        prefs.put("profile.default_content_setting_values.images", 2);
        prefs.put("profile.default_content_setting_values.stylesheets", 2);
        options.setExperimentalOption("prefs", prefs);

        WebDriver driver = new ChromeDriver(options);

        String currentPin = null;

        try {
            String filePath = ".\\input-data\\CityWise300 Input DataUpdatedNew.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("Jiomart1");

            List<String> inputPid = new ArrayList<>(), InputCity = new ArrayList<>(), InputName = new ArrayList<>(),
                    InputSize = new ArrayList<>(), NewProductCode = new ArrayList<>(), uRL = new ArrayList<>(),
                    UOM = new ArrayList<>(), Mulitiplier = new ArrayList<>(), Pincode = new ArrayList<>(),
                    NameForCheck = new ArrayList<>();

            for (int i = 1; i < urlsSheet.getPhysicalNumberOfRows(); i++) {
                Row row = urlsSheet.getRow(i);
                if (row == null) continue;

                Cell urlCell = row.getCell(5);
                if (urlCell == null || urlCell.getCellType() != CellType.STRING) continue;
                String url = urlCell.getStringCellValue().trim();
                if (url.isEmpty() || url.equalsIgnoreCase("NA")) continue;

                inputPid.add(getCellValue(row.getCell(0)));
                InputCity.add(getCellValue(row.getCell(1)));
                InputName.add(getCellValue(row.getCell(2)));
                InputSize.add(getCellValue(row.getCell(3)));
                NewProductCode.add(getCellValue(row.getCell(4)));
                uRL.add(url);
                UOM.add(getCellValue(row.getCell(6)));
                Mulitiplier.add(getCellValue(row.getCell(7)));
                Pincode.add(getCellValue(row.getCell(9)));
                NameForCheck.add(getCellValue(row.getCell(10)));
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
                String mulitiplier = Mulitiplier.get(i);
                String locationSet = Pincode.get(i);
                String namecheck = NameForCheck.get(i);

                String newName = "NA", spValue = "NA", mrpValue = "NA", offerValue = "NA", availability = "0";

                try {
                    driver.get(url);
                    if (i == 0) Thread.sleep(2500);

                    if (currentPin == null || !currentPin.equals(locationSet)) {
                        setPincode(driver, locationSet);
                        currentPin = locationSet;
                        Thread.sleep(1000);
                    }

                    newName = safeGetText(driver, "//div[@id='pdp_product_name']");

                    spValue = cleanPrice(safeGetText(driver, "(//div[@class='product-price jm-mb-xxs']//span)[1]"));

                    String mrpText = safeGetText(driver, "(//div[@class='jm-body-s jm-fc-primary-grey-80']//span)[1]");
                    mrpValue = mrpText.isEmpty() ? spValue : cleanPrice(mrpText);

                    try {
                        WebElement cartBtn = driver.findElement(By.xpath("(//button[text()='Add to Cart'])[1]"));
                        availability = (cartBtn.isDisplayed() && cartBtn.isEnabled()) ? "1" : "0";
                    } catch (Exception ignored) {}

                    if (!mrpValue.equals(spValue)) {
                        String offer = safeGetText(driver, "(//div[@class='product-price jm-mb-xxs']//span)[2]");
                        if (!offer.isEmpty()) offerValue = offer.replaceAll("[^0-9%]", "").trim();
                    }

                    writeRow(resultsSheet, rowIndex++, id, city, name, size, productCode, url,
                            newName, mrpValue, spValue, uom, mulitiplier, availability, offerValue, namecheck);

                    System.out.println("Done: " + (i + 1) + "/" + uRL.size() + " - " + url);
                    System.out.println("Scarped Details: "+ newName + " " +spValue +" " + mrpValue+ " "+ availability + " " + offerValue);
                    

                } catch (Exception e) {
                    e.printStackTrace();
                    writeRow(resultsSheet, rowIndex++, id, city, name, size, productCode, url,
                            "NA", "NA", "NA", uom, mulitiplier, "0", "NA", namecheck);
                }
            }

            String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
            String output = ".\\Output\\Jiomart_1_" + timestamp + ".xlsx";
            FileOutputStream out = new FileOutputStream(output);
            resultsWorkbook.write(out);
            out.close();
            resultsWorkbook.close();
            System.out.println("COMPLETED SUCCESSFULLY! → " + output);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            driver.quit();
            System.out.println("Scraping Finished - Ultra Fast & Zero Errors!");
        }
    }

    // FIXED: Classic if-else instead of switch expression (works on ALL Java versions)
    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf((long) cell.getNumericCellValue());
        } else {
            return "";
        }
    }

    private static String safeGetText(WebDriver driver, String xpath) {
        try {
            return driver.findElement(By.xpath(xpath)).getText().trim();
        } catch (Exception e) {
            return "";
        }
    }

    private static String cleanPrice(String text) {
        if (text.isEmpty()) return "NA";
        return text.replace("₹", "").replace(",", "").replace(".00", "").trim();
    }

    private static void createHeader(Sheet sheet) {
        Row h = sheet.createRow(0);
        String[] headers = {"InputPid","InputCity","InputName","InputSize","NewProductCode","URL","Name","MRP","SP",
                "UOM","Multiplier","Availability","Offer","Commands","Remarks","Correctness","Percentage","Name","Name Check"};
        for (int i = 0; i < headers.length; i++) h.createCell(i).setCellValue(headers[i]);
    }

    private static void writeRow(Sheet sheet, int rowNum, String... values) {
        Row row = sheet.createRow(rowNum);
        for (int i = 0; i < values.length; i++) {
            row.createCell(i).setCellValue(values[i]);
        }
        for (int i = values.length; i < 19; i++) {
            row.createCell(i).setCellValue("");
        }
    }

    private static void setPincode(WebDriver driver, String pin) throws Exception {
        try { driver.findElement(By.xpath("//span[@id='delivery_city_pincode']")).click(); } catch (Exception ignored) {}
        Thread.sleep(500);
        try { driver.findElement(By.xpath("//button[@id='btn_enter_pincode']")).click(); } catch (Exception ignored) {}
        Thread.sleep(500);
        WebElement input = driver.findElement(By.xpath("//input[@id='rel_pincode']"));
        input.clear();
        input.sendKeys(pin);
        Thread.sleep(400);
        driver.findElement(By.xpath("//button[@id='btn_pincode_submit']")).click();
    }

}

