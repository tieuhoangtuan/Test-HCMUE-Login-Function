package excelRead_Write;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class ReadWriteExcel {

	public static void main(String[] args) throws IOException, InterruptedException {
		
		// Đọc file TestCase.xlsx
		File file = new File(System.getProperty("user.dir") + "\\TestData\\" + "TestCase" + ".xlsx");
		FileInputStream inputstream=new FileInputStream(file);
		XSSFWorkbook wb=new XSSFWorkbook(inputstream);
		XSSFSheet sheet=wb.getSheet("LoginDetails");
		
		/* Mở trang web để đăng nhập HCMUE */
		WebDriverManager.chromedriver().setup();
		WebDriver driver=new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(20,TimeUnit.SECONDS);
		driver.get("https://online.hcmue.edu.vn/");
		driver.findElement(By.xpath("//a[text()='Đăng nhập']")).click();
		
		// Lưu giá trị vào trong biến  ( Trong trường hợp này là: Usernames và Passwords)
		XSSFRow row=null;
		XSSFCell cell=null;
		String userName=null;
		String password=null;
		
		for (int i=1; i<=sheet.getLastRowNum();i++)
		{
			row=sheet.getRow(i);
			for ( int j=1;j<row.getLastCellNum();j++)
			{
				cell=row.getCell(j);
				cell.setCellType(cell.CELL_TYPE_STRING);
				
				if(j==1) 
				{
					userName=cell.getStringCellValue();
				}
				if(j==2) 
				{
					password=cell.getStringCellValue();
				}				
			}
			// Truyền giá trị userName vào input Tên đăng nhập
			driver.findElement(By.id("ContentPlaceHolder1_ctl00_ctl00_txtUserName")).sendKeys(userName);
			// Truyền giá trị passWord vào input Mật mã
			driver.findElement(By.id("ContentPlaceHolder1_ctl00_ctl00_txtPassword")).sendKeys(password);
			// Click chuột vào nút Đăng nhập
			driver.findElement(By.id("ContentPlaceHolder1_ctl00_ctl00_btLogin")).click();
			
			// Kiểm tra xem Đăng nhập thành công hay thất bại
			String result=null;
			try 
			{	
				// Tìm trên web xem có tồn tại element nào có text = 'Đăng thoát'
				Boolean isLoggedIn=driver.findElement(By.xpath("//a[text()='Đăng Thoát']")).isDisplayed();
				
				// Nếu có --> Đăng nhập thành công!
				if(isLoggedIn==true)
				{
					result="PASS";
					// Viết kết quả PASS vào trong Excel
					cell=row.createCell(3);
					cell.setCellType(cell.CELL_TYPE_STRING);
					cell.setCellValue(result);
					
					
				}
				// In ra console kết quả đăng nhập
				System.out.println("User Name : " + userName + " ---- > " + "Password : "  + password + "-----> Login success ? ------> " + result);
				
				// Click chuột vào nút Đăng thoát
				driver.findElement(By.xpath("//a[text()='Đăng Thoát']")).click();
			}
			catch(Exception e)
			{
				// Tìm trên web xem có tồn tại element nào có text = 'Mật mã truy cập không chính xác.'
				Boolean isError=driver.findElement(By.xpath("//td[text()='Mật mã truy cập không chính xác.']")).isDisplayed();
				
				// Nếu có --> Đăng nhập thất bại
				if(isError==true)
				{
					result="FAIL";
					// Viết kết quả Fail vào trong Excel
					cell=row.createCell(3);
					cell.setCellType(cell.CELL_TYPE_STRING);
					cell.setCellValue(result);
				}
				// In ra console kết quả đăng nhập
				System.out.println("User Name : " + userName + " ---- > " + "Password : "  + password + "-----> Login success ? ------> " + result);
			}
			Thread.sleep(1000);
			driver.findElement(By.xpath("//a[text()='Đăng nhập']")).click();
		}
		FileOutputStream fos=new FileOutputStream(file);
		wb.write(fos);
		fos.close();
		driver.close();	
	}
}
