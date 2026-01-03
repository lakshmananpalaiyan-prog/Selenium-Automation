package scheduler;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;


public class Dmart {

	public static void main(String[] args) throws Exception {
		ChromeOptions options = new ChromeOptions();
	options.addArguments("--headless"); // Run Chrome in headless mode
	options.addArguments("--disable-gpu"); // Disable GPU acceleration
	options.addArguments("--window-size=1920,1080");   //Set window size to full HD
	options.addArguments("--start-maximized");	

	ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(1);

	// Schedule the task to run every day at 7:00 AM
	Calendar now = Calendar.getInstance();
	Calendar nextRunTime = Calendar.getInstance();
	nextRunTime.set(Calendar.HOUR_OF_DAY, 12);//4
	nextRunTime.set(Calendar.MINUTE, 50);
	nextRunTime.set(Calendar.SECOND, 0);

	long initialDelay = nextRunTime.getTimeInMillis() - now.getTimeInMillis();
	if (initialDelay < 0) {
		initialDelay += 24 * 60 * 60 * 1000; // If it's already past 7 AM, schedule for the next day
	}

	scheduler.scheduleAtFixedRate(() -> {
		try {
			System.out.println("Starting web scraping task...");
			Dmart_citywise.runWebScraping();
			System.out.println("Web scraping task completed.");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}, initialDelay, 24 * 60 * 60 * 1000, TimeUnit.MILLISECONDS);
}
public static void runWebScraping() throws Exception{

		ChromeOptions options = new ChromeOptions();
		WebDriver driver = new ChromeDriver(options);
		driver.manage().window().maximize();
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

		List<String> inputPid = new ArrayList<>(), InputCity = new ArrayList<>(), InputName = new ArrayList<>(),
				InputSize = new ArrayList<>(), NewProductCode = new ArrayList<>(), uRL = new ArrayList<>(),               
				Pincode = new ArrayList<>();

		Workbook resultsWorkbook = new XSSFWorkbook();
		Sheet resultsSheet = resultsWorkbook.createSheet("Results");
		createHeaderRow(resultsSheet);

		int rowIndex = 1;
		String currentPin = null;

		try (FileInputStream file = new FileInputStream("./input-data/CityWise300 Input DataUpdatedNew.xlsx");
				Workbook urlsWorkbook = new XSSFWorkbook(file)) {

			Sheet urlsSheet = urlsWorkbook.getSheet("Dmart");
			int rowCount = urlsSheet.getPhysicalNumberOfRows();

			for (int i = 1; i < rowCount; i++) {
				Row row = urlsSheet.getRow(i);
				if (row.getCell(5) != null) {
					String url = row.getCell(5).getCellType() == CellType.STRING
							? row.getCell(5).getStringCellValue()
									: String.valueOf(row.getCell(5).getNumericCellValue());

					inputPid.add(row.getCell(0).getStringCellValue());
					InputCity.add(row.getCell(1).getStringCellValue());
					InputName.add(row.getCell(2).getStringCellValue());
					InputSize.add(row.getCell(3).getStringCellValue());
					NewProductCode.add(row.getCell(4).getStringCellValue());
					uRL.add(url);
					Pincode.add(row.getCell(9).getStringCellValue());                    
				}
			}

			for (int i = 0; i < uRL.size(); i++) {
				String url = uRL.get(i);
				String locationSet = Pincode.get(i);
				if (url.isEmpty() || url.equalsIgnoreCase("NA")) {
					writeResults(resultsSheet, rowIndex++, inputPid.get(i), InputCity.get(i), InputName.get(i),
							InputSize.get(i), "NA", url, "NA", "NA", "NA",
							"NA", "NA", "NA", "NA");
					continue;
				}

				try {
					driver.get(url);
					Thread.sleep(2000);

					List<WebElement> unAvailable = driver.findElements(By.xpath("//div[text()='Currently Unavailable']"));
					if (!unAvailable.isEmpty()) {
						writeResults(resultsSheet, rowIndex++, inputPid.get(i), InputCity.get(i), InputName.get(i),
								InputSize.get(i), "NA", "NA", "NA", "NA", "NA",
								"NA", "NA", "NA", "NA");
						System.out.println("Product URL is unavailable: " + url);
						continue;  // Skip this URL and proceed with the next one
					}

					if (currentPin == null || !currentPin.equals(locationSet)) {
						driver.findElement(By.xpath("//div[@class='header_pincode__KryhE']")).click();
						Thread.sleep(1000);
						driver.findElement(By.xpath("//input[@id='pincodeInput']")).sendKeys(locationSet);
						Thread.sleep(3000);
						driver.findElement(By.xpath("(//div[@class='pincode-widget_pincode-right__TwcOu'])[1]"))
						.click();
						Thread.sleep(1000);
						driver.findElement(By.xpath("//button[.='CONFIRM LOCATION']")).click();
						Thread.sleep(2000);
						currentPin = locationSet;
						driver.get(url);
					}

					String productId = extractProductId(url);
					String newName = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h1"))).getText();

					String weight = extractWeight(driver, InputSize.get(i)); 
					String mrpValue = extractMRP(driver);
					String spValue = extractSP(driver);
					String offerValue = extractOffer(driver);

					String pageSource = driver.getPageSource();
					String availability = (pageSource.contains("Out Of Stock") || pageSource.contains("currently unavailable")) ? "0" : "1";

					String finalMultiplier = UOMUtils.calculateMultiplier(InputSize.get(i), weight, newName);

					writeResults(resultsSheet, rowIndex++, inputPid.get(i), InputCity.get(i), InputName.get(i),
							InputSize.get(i), productId, url, newName, mrpValue, spValue,
							weight, finalMultiplier, availability, offerValue);

					System.out.println("Data Extracting: "+ url);
					System.out.println(newName);
					System.out.println(mrpValue);
					System.out.println(spValue);
					System.out.println(weight);
					System.out.println(finalMultiplier);
					System.out.println(availability);
					System.out.println(offerValue);

				} catch (Exception e) {
					writeResults(resultsSheet, rowIndex++, inputPid.get(i), InputCity.get(i), InputName.get(i),
							InputSize.get(i), "NA", url, "NA", "NA", "NA",
							"NA", "NA", "NA", "NA");
					System.out.println("Error fetching data for URL: " + url);
					e.printStackTrace();
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			saveResultsToExcel(resultsWorkbook);
			if (driver != null) driver.quit();
		}
	}

	private static void createHeaderRow(Sheet sheet) {
		Row headerRow = sheet.createRow(0);
		String[] headers = {"InputPid", "InputCity", "InputName", "InputSize", "NewProductCode", "URL", "Name", "MRP",
				"SP", "UOM", "Multiplier", "Availability", "Offer"};
		for (int i = 0; i < headers.length; i++) {
			headerRow.createCell(i).setCellValue(headers[i]);
		}
	}

	private static void writeResults(Sheet sheet, int rowIndex, String... values) {
		Row row = sheet.createRow(rowIndex);
		for (int i = 0; i < values.length; i++) {
			row.createCell(i).setCellValue(values[i]);
		}
	}

	private static void saveResultsToExcel(Workbook workbook) {
		try {
			String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
			try (FileOutputStream out = new FileOutputStream("./Output/DMART_OutputData_" + timestamp + ".xlsx"))
			{
				workbook.write(out);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static String extractProductId(String url) {
		try {
			int startIndex = url.indexOf("selectedProd=") + 13;
			int endIndex = url.indexOf("&", startIndex);
			return url.substring(startIndex, endIndex == -1 ? url.length() : endIndex);
		} catch (Exception e) {
			return "NA";
		}
	}

	private static String extractMRP(WebDriver driver) {
		try {
			List<WebElement> elements = driver.findElements(By.xpath("//span[contains(., 'MRP')]"));
			if (elements.isEmpty()) return "NA";
			String text = elements.get(0).getText().replaceAll("[₹]", "").replaceAll("(?i)MRP", "").trim();
			return text;
		} catch (Exception e) {
			return "NA";
		}
	}

	private static String extractSP(WebDriver driver) {
		try {
			List<WebElement> elements = driver.findElements(By.xpath("(//span[@class='price-details-component_value__IvVER'])[1]"));
			return elements.isEmpty() ? "NA" : elements.get(0).getText().replace("₹", "").trim();
		} catch (Exception e) {
			return "NA";
		}
	}

	private static String extractOffer(WebDriver driver) {
		try {
			List<WebElement> elements = driver.findElements(By.xpath("//div[@class='price-details-component_saveHighlighter__FIIS_']//div"));
			return elements.isEmpty() ? "NA" : elements.get(0).getText().trim();
		} catch (Exception e) {
			return "NA";
		}
	}
	
	private static String extractWeight(WebDriver driver, String inputUOM) {
	    try {
	        String text = driver.findElement(By.xpath("//h1")).getText();
	        String bracketUOM = UOMUtils.extractBracketUOM(text);
	        String colonUOM = null;

	        if (text.contains(":")) {
	            colonUOM = text.split(":")[1].trim();
	        }

	        boolean inputIsVolume = UOMUtils.isVolumeUnit(inputUOM);
	        boolean inputIsWeight = UOMUtils.isWeightUnit(inputUOM);

	        // Prefer bracket UOM if it matches input type
	        if (bracketUOM != null && !bracketUOM.equals("NA")) {
	            if ((inputIsVolume && UOMUtils.isVolumeUnit(bracketUOM)) ||
	                (inputIsWeight && UOMUtils.isWeightUnit(bracketUOM))) {
	                return bracketUOM;
	            }
	        }

	        // If bracket UOM not suitable, fallback to colon UOM
	        if (colonUOM != null && !colonUOM.isEmpty()) {
	            return colonUOM;
	        }

	        return "NA";
	    } catch (Exception e) {
	        return "NA";
	    }
	}

	
}

class UOMUtils {

		public static double normalizeToBaseUnit(String uom) {
		if (uom == null || uom.trim().isEmpty()) return 0;

		/*
		uom = uom.toLowerCase().replaceAll("\\s+", "");
		// Normalize plurals and common variations
		uom = uom.replace("kgs", "kg").replace("kilograms", "kg").replace("kilogram", "kg")
				.replace("grams", "g").replace("gram", "g").replace("gm", "g")
				.replace("litres", "l").replace("liters", "l").replace("liter", "l")
				.replace("ltr", "l").replace("lts", "l")
				.replace("milliliters", "ml").replace("millilitre", "ml").replace("mls", "ml");
				*/
		
		uom = uom.toLowerCase().replaceAll("\\s+", ""); // remove all whitespace

		// multiplier for 1000 
		uom = uom.replaceAll("kilograms?|kgs?", "kg")
		         .replaceAll("grams?|gms?|gm", "g")
		         .replaceAll("litres?|liters?|litre|liter|ltr|lts?", "l")
		         .replaceAll("millilitres?|milliliters?|mls?", "ml");

		
		try {
			if (uom.endsWith("kg")) return Double.parseDouble(uom.replace("kg", "")) * 1000;
			if (uom.endsWith("g")) return Double.parseDouble(uom.replace("g", ""));
			if (uom.endsWith("l")) return Double.parseDouble(uom.replace("l", "")) * 1000;
			if (uom.endsWith("ml")) return Double.parseDouble(uom.replace("ml", ""));
			return Double.parseDouble(uom.replaceAll("[^\\d.]", ""));
		} catch (Exception e) {
			return 0;
		}
	}

	// Extract the unit inside brackets from scraped product name, e.g., "(1 L)" → "1 L"
	public static String extractBracketUOM(String name) {
		if (name == null || !name.contains("(") || !name.contains(")")) return "NA";
		try {
			return name.substring(name.indexOf("(") + 1, name.indexOf(")")).trim();
		} catch (Exception e) {
			return "NA";
		}
	}

	// Detect if a UOM string represents volume units
	public static boolean isVolumeUnit(String uom) {
		if (uom == null) return false;
		return uom.toLowerCase().matches(".*\\b(l|ltr|litre|liters|liter|ml|milliliter|millilitre)\\b.*");
	}

	// Detect if a UOM string represents weight units
	public static boolean isWeightUnit(String uom) {
		if (uom == null) return false;
		return uom.toLowerCase().matches(".*\\b(g|gram|kg|kgs|kilogram|kilograms|gm)\\b.*");
	}

	public static String calculateMultiplier(String inputUOM, String extractedUOM, String scrapedName) {
	    double inputVal = normalizeToBaseUnit(inputUOM);
	    double extractedVal = normalizeToBaseUnit(extractedUOM);

	    boolean inputIsVolume = isVolumeUnit(inputUOM);
	    boolean extractedIsWeight = isWeightUnit(extractedUOM);

	    // If input is volume but extracted is weight (likely mismatch), fallback to bracket UOM
	    if (inputIsVolume && extractedIsWeight) {
	        String bracketUOM = extractBracketUOM(scrapedName);
	        extractedVal = normalizeToBaseUnit(bracketUOM);
	    }
	    
	    // If input and extracted UOM are same and values are equal, multiplier = 1
	    if (Math.abs(inputVal - extractedVal) < 0.01) {
	        return "1.00";
	    }

	    if (inputVal == 0 || extractedVal == 0) {
	        System.out.println("Multiplier calculation failed: input=" + inputUOM + ", extracted=" + extractedUOM + ", bracket=" + extractBracketUOM(scrapedName));
	        return "NA";
	    }

	    double multiplier = inputVal / extractedVal;
	    
//	    // round the multiplier value eg. 1.04 means 1 or 1.05 means 1.1
//	    double roundedMultiplier;
//	    if (multiplier - Math.floor(multiplier) < 0.05) {
//	        roundedMultiplier = Math.floor(multiplier * 10) / 10.0;  // round down to nearest 0.1
//	    } else {
//	        roundedMultiplier = Math.ceil(multiplier * 10) / 10.0;   // round up to nearest 0.1
//	    }
	    return String.format("%.2f", multiplier);

	}

}


