import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.WebDriverWait;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Flipkart {
    public static void main(String[] args) throws Exception {
        ChromeOptions options = new ChromeOptions();
        // options.addArguments("--headless=new"); // Uncomment for headless
        options.addArguments("--disable-gpu", "--no-sandbox", "--disable-dev-shm-usage");
        options.addArguments("--disable-extensions", "--disable-images", "--blink-settings=imagesEnabled=false");
        options.setPageLoadStrategy(PageLoadStrategy.EAGER);
        WebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));

        // CHANGE THIS PINCODE AS PER CITY
        String TARGET_PINCODE = "700001"; // Kolkata pincode (you can change to 110001, 400001, etc.)

        try {
            // STEP 1: Set Pincode ONCE using Flipkart's homepage
            setPincodeOnce(driver, wait, TARGET_PINCODE);
            System.out.println("Pincode " + TARGET_PINCODE + " set successfully for the entire session!");

            // STEP 2: Now read Excel and start scraping
            List<String> inputPid = new ArrayList<>(), InputCity = new ArrayList<>(), InputName = new ArrayList<>(), InputSize = new ArrayList<>(),
                    NewProductCode = new ArrayList<>(), uRL = new ArrayList<>(), UOM = new ArrayList<>(), Mulitiplier = new ArrayList<>(),
                    Availability = new ArrayList<>(), Pincode = new ArrayList<>(), NameForCheck = new ArrayList<>();

            FileInputStream file = new FileInputStream(".\\input-data\\CityWise300 Input DataUpdatedNew.xlsx");
            Workbook wb = new XSSFWorkbook(file);
            Sheet sheet = wb.getSheet("FK_Kol");

            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                inputPid.add(getCellValue(row.getCell(0)));
                InputCity.add(getCellValue(row.getCell(1)));
                InputName.add(getCellValue(row.getCell(2)));
                InputSize.add(getCellValue(row.getCell(3)));
                NewProductCode.add(getCellValue(row.getCell(4)));
                uRL.add(getCellValue(row.getCell(5)));
                UOM.add(getCellValue(row.getCell(6)));
                Mulitiplier.add(getCellValue(row.getCell(7)));
                Availability.add(getCellValue(row.getCell(8)));
                Pincode.add(getCellValue(row.getCell(9)));
                NameForCheck.add(getCellValue(row.getCell(10)));
            }
            wb.close();
            file.close();

            // Output Excel Setup
            Workbook resultWb = new XSSFWorkbook();
            Sheet resultSheet = resultWb.createSheet("Results");
            String[] headers = {"InputPid", "InputCity", "InputName", "InputSize", "NewProductCode", "URL", "Name", "MRP","SP","UOM","Multiplier","Availability","Offer","Commands","Remarks", "Correctness","Percentage","Name","Name Check"};
            Row headerRow = resultSheet.createRow(0);
            for (int i = 0; i < headers.length; i++) headerRow.createCell(i).setCellValue(headers[i]);

            int rowIndex = 1;
            System.out.println("Starting scraping " + uRL.size() + " products with pincode " + TARGET_PINCODE + "...");

            for (int i = 0; i < uRL.size(); i++) {
                String url = uRL.get(i);
                String id = inputPid.get(i);
                String newName = "NA", mrpValue = "NA", spValue = "NA", offerValue = "NA", availabilityStatus = "NA";
                String scrapedUOM = "NA"; // Declare outside to fix scope error

                if (!url.isEmpty() && !url.equalsIgnoreCase("NA")) {
                    try {
                        driver.get(url);
                        Thread.sleep(800); // Small buffer for dynamic content

                        // Name
                        try {
                            newName = wait.until(d -> d.findElement(By.cssSelector("h1 span, span.B_NuCI"))).getText().trim();
                        } catch (Exception ignored) {}

                        // SP
                        try {
                            spValue = driver.findElement(By.xpath("//div[@class='hZ3P6w bnqy13']"))
                                    .getText().replace("₹", "").replace(",", "");
                        } catch (Exception ignored) {
                            spValue = "NA";
                        }

                        // MRP
                        try {
                            String mrpText = driver.findElement(By.xpath("//div[@class='kRYCnD yHYOcc']"))
                                    .getText().replace("₹", "").replace(",", "");
                            mrpValue = mrpText.isEmpty() ? spValue : mrpText;
                        } catch (Exception ignored) {
                            mrpValue = spValue;
                        }

                        // Offer
                        if (!mrpValue.equals(spValue) && !"NA".equals(spValue)) {
                            try {
                                offerValue = driver.findElement(By.xpath("//div[@class='HQe8jr rASMtN']"))
                                        .getText().replace("off", "Off");
                            } catch (Exception ignored) {}
                        } else {
                            offerValue = "NA";
                        }

                        // UOM (from selected variant selector)
                        try {
                            scrapedUOM = driver.findElement(By.xpath("//div[@data-testid='sellerQuantity']//div[contains(@class, '*3I1e3u')]")).getText().trim();
                            if (scrapedUOM.isEmpty()) scrapedUOM = "NA";
                        } catch (Exception ignored) {
                            // Fallback: Extract from product name (last parentheses)
                            try {
                                int startIdx = newName.lastIndexOf("(");
                                int endIdx = newName.lastIndexOf(")");
                                if (startIdx != -1 && endIdx != -1 && startIdx < endIdx) {
                                    scrapedUOM = newName.substring(startIdx + 1, endIdx).trim();
                                }
                            } catch (Exception e) {
                                scrapedUOM = UOM.get(i); // Use input UOM as last resort
                            }
                        }

                        // Availability
                        availabilityStatus = "1";
                        String pageSource = driver.getPageSource().toLowerCase();
                        if (pageSource.contains("currently unavailable") || pageSource.contains("sold out") || pageSource.contains("out of stock in this area")) {
                            availabilityStatus = "0";
                        }

                        System.out.printf("Done [%d/%d] %s → SP: %s | MRP: %s | In Stock: %s%n", i + 1, uRL.size(), id, spValue, mrpValue, availabilityStatus);
                    } catch (Exception e) {
                        System.out.println("Failed: " + url);
                        newName = "NA"; mrpValue = "NA"; spValue = "NA"; offerValue = "NA"; availabilityStatus = "NA";
                    }
                } else {
                    newName = "NA"; mrpValue = "NA"; spValue = "NA"; offerValue = "NA";
                }

                // Calculate multiplier (using scrapedUOM for accuracy; fallback to input if NA)
                String uomForCalc = "NA".equals(scrapedUOM) ? UOM.get(i) : scrapedUOM;
                String calculatedMultiplier = calculateMultiplier(InputSize.get(i), uomForCalc);

                // Write row
                Row row = resultSheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(id);
                row.createCell(1).setCellValue(InputCity.get(i));
                row.createCell(2).setCellValue(InputName.get(i));
                row.createCell(3).setCellValue(InputSize.get(i));
                row.createCell(4).setCellValue(NewProductCode.get(i));
                row.createCell(5).setCellValue(url);
                row.createCell(6).setCellValue(newName);
                row.createCell(7).setCellValue(mrpValue);
                row.createCell(8).setCellValue(spValue);
                row.createCell(9).setCellValue(scrapedUOM);
                row.createCell(10).setCellValue(calculatedMultiplier);
                row.createCell(11).setCellValue(availabilityStatus);
                row.createCell(12).setCellValue(offerValue);
                row.createCell(13).setCellValue("NA"); // Commands
                row.createCell(14).setCellValue("NA"); // Remarks
                row.createCell(15).setCellValue("NA"); // Correctness
                row.createCell(16).setCellValue("NA"); // Percentage
                row.createCell(17).setCellValue(NameForCheck.get(i)); // Name Check (fixed index)
            }

            // Save output
            String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
            String outputPath = ".\\Output\\Flipkart_Kol" + timestamp + ".xlsx";
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                resultWb.write(fos);
            }
            resultWb.close();
            System.out.println("Scraping completed! Saved: " + outputPath);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            driver.quit();
        }
    }

    // SET PINCODE ONLY ONCE – MOST RELIABLE METHOD
    private static void setPincodeOnce(WebDriver driver, WebDriverWait wait, String pincode) throws InterruptedException {
        driver.get("https://www.flipkart.com/mivi-play-12hrs-playback-bass-boosted-tws-feature-ipx4-5-w-portable-bluetooth-speaker/p/itm9dcdfcd2431db");
        Thread.sleep(2000);
        try {
            // Close login popup if appears
            driver.findElement(By.xpath("//button[contains(text(),'✕')]")).click();
            Thread.sleep(1000);
        } catch (Exception ignored) {}
        try {
            WebElement pinBox = wait.until(d -> d.findElement(By.xpath("//input[@placeholder='Enter delivery pincode' or @class='AFOXgu' or contains(@class,'pincode')]")));
            pinBox.clear();
            pinBox.sendKeys(pincode);
            WebElement checkBtn = driver.findElement(By.xpath("//span[text()='Check' or contains(@class,'i40dM4')]//parent::button | //button[.//span[text()='Check']]"));
            checkBtn.click();
            Thread.sleep(2000); // Wait for pincode to apply
            // Optional: Verify pincode is set
            String currentPin = driver.findElement(By.xpath("//div[contains(text(),'Deliver to')]//following-sibling::div//div[contains(@class,'_1XFP')]")).getText();
            System.out.println("Pincode applied: " + currentPin);
        } catch (Exception e) {
            System.out.println("Could not set pincode automatically. Please set it manually once and re-run.");
        }
    }

    public static String calculateMultiplier(String inputSize, String uom) {
        try {
            if (uom.toLowerCase().contains("pack of 1")) {
                return "1";
            }
            double inputSizeValue = parseSizeToGramsOrML(inputSize);
            double uomValue = parseSizeToGramsOrML(uom);
            if (uomValue == 0) return "NA";
            double multiplier = inputSizeValue / uomValue;
            double roundedUp = Math.ceil(multiplier * 10.0) / 10.0;
            return String.format("%.1f", roundedUp);
        } catch (Exception e) {
            return "NA";
        }
    }

    public static double parseSizeToGramsOrML(String sizeStr) {
        if (sizeStr == null || sizeStr.trim().isEmpty()) return 0;
        sizeStr = sizeStr.toLowerCase().replace("pack", "").replaceAll("[()\s]", "");
        String[] parts = sizeStr.split("x");
        double total = 1.0;
        for (String part : parts) {
            total *= parseSingleUnit(part);
        }
        return total;
    }

    public static double parseSingleUnit(String unitStr) {
        unitStr = unitStr.trim().toLowerCase();
        String unit = "";
        if (unitStr.contains("kg")) unit = "kg";
        else if (unitStr.contains("ltr") || unitStr.contains("l")) unit = "l";
        else if (unitStr.contains("g")) unit = "g";
        else if (unitStr.contains("ml")) unit = "ml";

        String numPart = unitStr.replaceAll("[a-z]+", "");
        double num;
        if (numPart.contains("-")) {
            String minStr = numPart.split("-")[0].replaceAll("[^\\d.]", "");
            num = minStr.isEmpty() ? 0 : Double.parseDouble(minStr);
        } else {
            String numStr = numPart.replaceAll("[^\\d.]", "");
            num = numStr.isEmpty() ? 0 : Double.parseDouble(numStr);
        }

        if (unit.equals("kg") || unit.equals("l")) {
            return num * 1000;
        } else if (unit.equals("g") || unit.equals("ml")) {
            return num;
        } else {
            return num; // For cases without unit, like "1"
        }
    }

    // Helper to safely get cell value
    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue().trim();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf((long) cell.getNumericCellValue());
        } else {
            return "";
        }
    }

}

