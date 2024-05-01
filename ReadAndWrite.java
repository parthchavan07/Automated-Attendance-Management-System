import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Date;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
class SeleniumJava {
    public void Meth1() throws IOException, InterruptedException {
    	Sort b = new Sort();
    	String[] NewData = b.Meth3();
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.get("http://127.0.0.1:5500/TableAbsent.html");
        WebElement table = driver.findElement(By.xpath("//table[@class='StudentsAttendance']"));
        java.util.List<WebElement> nameCells = table.findElements(By.xpath("//tbody/tr/td[1]"));
        java.util.List<WebElement> radioButtons = table.findElements(By.xpath("//tbody/tr/td[2]/form/label/input"));
        for (int i = 0; i < nameCells.size(); i++) {
            WebElement nameCell = nameCells.get(i);
            String studentName = nameCell.getText().trim();
            boolean isAbsent = false;
            for (String absentName : NewData) {
                if (studentName.contains(absentName)){
                    isAbsent = true;
                    break;
                }
            }
            if (isAbsent){
                radioButtons.get(i * 2 + 1).click();
            } else {
                radioButtons.get(i * 2).click(); 
            }         Thread.sleep(700);
        }
        Thread.sleep(1200);
        driver.findElement(By.xpath("//input[@class='OnWoc']")).click();
        Date TodayDate=new Date();
       String SSFileName= TodayDate.toString().replace(" ", "-").replace(":", "-");
        File screenshotFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
        FileUtils.copyFile(screenshotFile, new File(".//screemshot//"+SSFileName+".png"));
        Thread.sleep(2000);
        driver.close();
    }
        }  
class Sort
{
	public String[] Meth3() throws IOException {
		ToRead a = new ToRead();
        String[][] data = a.Meth2();
        int absentCount = 0;
        for (int i = 0; i < data.length; i++) {
            if ("A".equalsIgnoreCase(data[i][2])) {
                absentCount++;
            }
        }
        String[] absentNames = new String[absentCount];
        int index = 0;
        for (int i = 0; i < data.length; i++) {
            if ("A".equalsIgnoreCase(data[i][2])) {
                absentNames[index] = data[i][0];
                index++;
            }  
        }
        return absentNames;
    }  
}
class ToRead
{
	public String[][] Meth2() throws IOException {
        String excelFilePath = ".\\DataFile\\Book2.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheet("sheet1");
        int rows = sheet.getLastRowNum() + 1;
        int cols = sheet.getRow(0).getLastCellNum();
        String[][] data = new String[rows][cols];
        int x=0,y=0;
        for (int r = 1; r < rows; r++) {
            Row row = sheet.getRow(r);
            for (int c = 0; c < cols ; c++) {
                Cell cell = row.getCell(c);
                CellType cellType = cell.getCellType();
                switch (cellType) {
                    case STRING:
                        data[x][y] = cell.getStringCellValue();
                        break;
                    case NUMERIC:
                    	data[x][y] = String.valueOf((int) cell.getNumericCellValue());
                         break;
                    case BOOLEAN:
                        data[x][y] = String.valueOf(cell.getBooleanCellValue());
                        break;
                    default:
                        data[x][y] = "";
                        System.out.println("Unsupported cell type at row " + r + ", column " + c);    
                }
                y+=1;
            }
            y=0;
            x+=1;
        }
        workbook.close();
        inputStream.close();
        return data;
}
}
public class ReadAndWrite {
    public static void main(String[] args) throws IOException, InterruptedException {
        SeleniumJava obj = new SeleniumJava();
        obj.Meth1();   
    }
}