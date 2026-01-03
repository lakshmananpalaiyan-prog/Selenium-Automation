package citywise;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.*;
import java.util.concurrent.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.chrome.*;
import org.openqa.selenium.support.ui.*;

public class SwiggyCityWise {

    public static void main(String[] args) throws Exception {

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--incognito");

        ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(1);

        Calendar now = Calendar.getInstance();
        Calendar nextRunTime = Calendar.getInstance();
        nextRunTime.set(Calendar.HOUR_OF_DAY, 12);
        nextRunTime.set(Calendar.MINUTE, 41);
        nextRunTime.set(Calendar.SECOND, 0);

        long initialDelay = nextRunTime.getTimeInMillis() - now.getTimeInMillis();
        if (initialDelay < 0) {
            initialDelay += 24 * 60 * 60 * 1000;
        }

        scheduler.scheduleAtFixedRate(() -> {
            try {
                System.out.println("Starting task...");
                runWebScraping(options);
                System.out.println("Completed.");
            } catch (Exception e) {
                e.printStackTrace();
            }
        }, initialDelay, 24 * 60 * 60 * 1000, TimeUnit.MILLISECONDS);
    }

    public static void runWebScraping(ChromeOptions options) throws Exception {

        WebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));

        try {

            String filePath = ".\\input-data\\swiggy_CitywiserunDaily.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook wb = new XSSFWorkbook(file);
            Sheet sheet = wb.getSheet("Mumbai");

            List<String> pid = new ArrayList<>(), city = new ArrayList<>(), name = new ArrayList<>(),
                    size = new ArrayList<>(), product = new ArrayList<>(), url = new ArrayList<>(),
                    uom = new ArrayList<>(), multi = new ArrayList<>(), avail = new ArrayList<>(),
                    pincode = new ArrayList<>();

            DataFormatter formatter = new DataFormatter();

            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {

                Row row = sheet.getRow(i);
                if (row == null || i == 0) continue;

                pid.add(getStringValue(row.getCell(0), formatter));
                city.add(getStringValue(row.getCell(1), formatter));
                name.add(getStringValue(row.getCell(2), formatter));
                size.add(getStringValue(row.getCell(3), formatter));
                product.add(getStringValue(row.getCell(4), formatter));
                url.add(getStringValue(row.getCell(5), formatter));
                uom.add(getStringValue(row.getCell(6), formatter));
                multi.add(getStringValue(row.getCell(7), formatter));
                avail.add(getStringValue(row.getCell(8), formatter));
                pincode.add(getStringValue(row.getCell(9), formatter));
            }

            Workbook outWB = new XSSFWorkbook();
            Sheet outSheet = outWB.createSheet("Results");

            String[] titles = {"InputPid", "InputCity", "InputName", "InputSize", "ProductID", "URL", "Name",
                    "MRP", "SP", "UOM", "Multiplier", "Availability", "Offer"};

            Row header = outSheet.createRow(0);
            for (int i = 0; i < titles.length; i++)
                header.createCell(i).setCellValue(titles[i]);

            int outRow = 1;
            int headerCount = 0;

            for (int i = 0; i < url.size(); i++) {

                String finalName = "NA", finalMRP = "NA", finalSP = "NA", finalUOM = "NA",
                        finalOffer = "NA", finalAvailability = "1";
                String webUom = "NA";
                double multiplier = 0.0;

                try {

                    if (i == 0) {
                        driver.get("https://www.swiggy.com/restaurants");

                        WebElement pinBtn = wait.until(ExpectedConditions
                                .presenceOfElementLocated(By.xpath("//span[.='Other']//span")));
                        pinBtn.click();

                        WebElement pinBox = wait.until(ExpectedConditions
                                .presenceOfElementLocated(By.xpath("//input[contains(@placeholder,'Search for area')]")));
                        pinBox.sendKeys(pincode.get(i));

                        Thread.sleep(1000);
                        wait.until(ExpectedConditions.elementToBeClickable(
                                By.xpath("(//div[@class='_2RwM6'])[1]"))).click();

                        Thread.sleep(1500);
                    }

                    driver.manage().window().maximize();
                    driver.get(url.get(i));
                    Thread.sleep(1500);

                    boolean popupStillThere = false;

                    try {
                        WebElement wrong = driver.findElement(By.xpath("//div[text()='Something went wrong!']"));
                        if (wrong.isDisplayed()) {

                            WebElement tryAgain = driver.findElement(
                                    By.xpath("//div[@data-testid='error-button']//button"));

                            for (int t = 0; t < 5; t++) {
                                tryAgain.click();
                                Thread.sleep(700);

                                try {
                                    wrong = driver.findElement(By.xpath("//div[text()='Something went wrong!']"));
                                    if (!wrong.isDisplayed()) break;
                                } catch (NoSuchElementException ne) {
                                    break;
                                }

                                tryAgain = driver.findElement(
                                        By.xpath("//div[@data-testid='error-button']//button"));
                            }
                        }
                    } catch (Exception ignored) {}

                    try {
                        WebElement stillWrong = driver.findElement(By.xpath("//div[text()='Something went wrong!']"));
                        popupStillThere = stillWrong.isDisplayed();
                    } catch (Exception ok) {
                        popupStillThere = false;
                    }

                    if (!popupStillThere) {

                        try {
                            finalName = driver.findElement(By.xpath("//div[@data-testid='item-name']")).getText();
                        } catch (Exception e) {
                            try {
                                finalName = driver.findElement(By.xpath("//h1")).getText();
                            } catch (Exception ex) {
                                finalName = "NA";
                            }
                        }

                        try {
                            finalSP = driver.findElement(By.xpath("//div[contains(@class,'_1bWTz')]"))
                                    .getText().replace("₹", "").trim();
                        } catch (Exception e) {
                            finalSP = "NA";
                        }

                        try {
                            finalMRP = driver.findElement(By.xpath("//div[contains(@class,'_2KTMQ')]"))
                                    .getText().replace("₹", "").trim();
                        } catch (Exception e) {
                            finalMRP = finalSP;
                        }

                        try {
                            webUom = driver.findElement(By.xpath(
                                            "//div[@class='_30iun']//div[@class='sc-gEvEer ymEfJ _11EdJ']"))
                                    .getText();
                        } catch (Exception e) {
                            try {
                                webUom = driver.findElement(By.xpath(
                                                "//div[@class='_2M_kP']//div[@class='sc-eqUAAy dEjugH _1TwvP']"))
                                        .getText();
                            } catch (Exception a) {
                                webUom = "NA";
                            }
                        }

                        multiplier = calculateMultiplier(size.get(i), webUom);

                        try {
                            finalOffer = driver.findElement(By.xpath(
                                            "//div[contains(@class,'sc-gEvEer bsYAwc _1WaLo')]"))
                                    .getText();
                            if (finalOffer.isEmpty()) finalOffer = "NA";
                        } catch (Exception e) {
                            finalOffer = "NA";
                        }

                        String src = driver.getPageSource();
                        if (src.contains("Sold Out") || src.contains("Unavailable") ||
                                src.contains("Currently Unavailable") || src.contains("out of stock"))
                            finalAvailability = "0";

                        headerCount++;
                        System.out.println("headercount = " + headerCount);
                        System.out.println("Data extracting:" + url.get(i));
                        System.out.println(finalName);
                        System.out.println(finalMRP);
                        System.out.println(finalSP);
                        System.out.println(webUom);
                        System.out.println("Multiplier: " + multiplier);
                        System.out.println(finalAvailability);
                        System.out.println(finalOffer);
                    }

                } catch (Exception e) {
                    System.out.println("Failed to extract : " + url.get(i));
                }

                Row r = outSheet.createRow(outRow++);
                r.createCell(0).setCellValue(pid.get(i));
                r.createCell(1).setCellValue(city.get(i));
                r.createCell(2).setCellValue(name.get(i));
                r.createCell(3).setCellValue(size.get(i));
                r.createCell(4).setCellValue(product.get(i));
                r.createCell(5).setCellValue(url.get(i));
                r.createCell(6).setCellValue(finalName);
                r.createCell(7).setCellValue(finalMRP);
                r.createCell(8).setCellValue(finalSP);
                r.createCell(9).setCellValue(webUom);
                r.createCell(10).setCellValue(multiplier);
                r.createCell(11).setCellValue(finalAvailability);
                r.createCell(12).setCellValue(finalOffer);
            }

            SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd_HHmmss");
            String output = ".\\Output\\Swiggy_OutputData_mum_" + df.format(new Date()) + ".xlsx";

            FileOutputStream fos = new FileOutputStream(output);
            outWB.write(fos);
            fos.close();

            System.out.println("File saved → " + output);

        } finally {
            driver.quit();
        }
    }

    /** SAFE CELL → STRING **/
    private static String getStringValue(Cell cell, DataFormatter formatter) {
        if (cell == null) return "";
        return formatter.formatCellValue(cell);   // prevents STRING vs NUMERIC crash
    }

    // ================= MULTIPLIER METHODS (UNCHANGED) =================

    private static double calculateMultiplier(String inputUom, String outputUom) {
        try {
            String input = inputUom.toLowerCase().trim();
            String output = outputUom.toLowerCase().trim();

            if (output.matches("\\d+\\s*pack.*")) {
                int inputPackCount = extractPackCount(input);
                int outputPackCount = extractPackCount(output);

                if (inputPackCount > 0 && outputPackCount > 0) {
                    if (inputPackCount == outputPackCount) {
                        return 1.0;
                    }
                }
            }

            double inputTotal = calculateTotalFromUom(input);
            double outputTotal = calculateTotalFromUom(output);

            if (outputTotal == 0) return 0;

            if (output.contains("pack") && !output.matches(".*\\(.*\\).*")) {
                if (input.contains("pack")) {
                    int inputPackCount = extractPackCount(input);
                    int outputPackCount = extractPackCount(output);
                    if (inputPackCount > 0 && outputPackCount > 0 && inputPackCount == outputPackCount) {
                        return 1.0;
                    }
                }
            }

            double multiplier = inputTotal / outputTotal;
            return Math.round(multiplier * 100.0) / 100.0;

        } catch (Exception e) {
            return 0;
        }
    }

    private static int extractPackCount(String text) {
        try {
            Matcher m = Pattern.compile("(\\d+)\\s*pack").matcher(text);
            if (m.find()) return Integer.parseInt(m.group(1));
        } catch (Exception ignored) {}
        return 0;
    }

    private static double convertToGrams(String qty, String unit) {
        double quantity = Double.parseDouble(qty);
        switch (unit.toLowerCase()) {
            case "kg": return quantity * 1000;
            case "g": return quantity;
            default: return quantity;
        }
    }

    private static double calculateTotalFromUom(String uom) {
        try {
            String[] parts = uom.split("[x*]");
            double total = 1.0;

            for (String part : parts) {
                part = part.trim();
                if (part.isEmpty()) continue;

                Matcher m = Pattern.compile("(\\d+\\.?\\d*)\\s*([a-zA-Z]*)").matcher(part);
                if (m.find()) {
                    String qty = m.group(1);
                    String unit = m.group(2);
                    if (unit.isEmpty()) unit = "g";
                    total *= convertToGrams(qty, unit);
                }
            }
            return total;

        } catch (Exception e) {
            return 0;
        }
    }
}

