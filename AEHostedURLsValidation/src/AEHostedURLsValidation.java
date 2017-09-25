import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.NoSuchElementException;
import java.util.Set;
import java.util.concurrent.TimeUnit;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.google.common.base.Function;


public class AEHostedURLsValidation {
	
	

		public static void main(String[] args) throws InvalidFormatException, IOException, InterruptedException {
			
			System.setProperty("webdriver.chrome.driver", "C:/SCS/Ananth/AEHostedURLsValidation/Driver/chromedriver.exe");
			WebDriver driver = new ChromeDriver();
			
			FileInputStream fis = new FileInputStream("C:/SCS/Ananth/AEHostedURLsValidation/HostedURLs.xlsx");
			Workbook w = WorkbookFactory.create(fis);
			Sheet s = w.getSheet("AEProductInstances");
			
			
			for(int row = 1; row <= s.getLastRowNum(); row++)
			{
			   String url = s.getRow(row).getCell(0).getStringCellValue();
				//To Pass Client URL from Excel
			   driver.get(url);
			   driver.manage().window().maximize();
			   
			   WebDriverWait wait = new WebDriverWait(driver, 120);

			    wait.until(new ExpectedCondition<Boolean>() {
			        public Boolean apply(WebDriver driver) {
			            return ((JavascriptExecutor) driver).executeScript(
			                "return document.readyState"
			            ).equals("complete");
			        }
			    });
			   
			   try
			   {
			   
			   if(driver.findElement(By.id("ddlLoginMethod")).isDisplayed())
			   {
				   System.out.println("Client having the ADFS for the Row " + row);
			   continue;
			   } 
			   } catch (org.openqa.selenium.NoSuchElementException e1)
			   
			   {
				   
				String name = s.getRow(row).getCell(1).getStringCellValue();
				//To Pass User Name from Excel
				try
				{
			        Select droplist = new Select(driver.findElement(By.id("drpUserName")));   
			        droplist.selectByVisibleText(name);

			        } catch (org.openqa.selenium.NoSuchElementException e2)
			        
			        {
			        
			        driver.findElement(By.id("txtUserName")).sendKeys(name);
			       }
			   
			   String pwd = s.getRow(row).getCell(2).getStringCellValue();
			   //To Pass Password from Excel
			   driver.findElement(By.id("txtPassword")).sendKeys(pwd);
			   
			   driver.findElement(By.id("cmdLogin")).click();
			   Thread.sleep(1500);
			   
			   try
			   {
			   
			   if(driver.findElement(By.id("lblErrmessage")).isDisplayed())
			   {
				   System.out.println("Invalid User Name or Password for the Row " + row);
			   continue;
			   } 
			   } catch (org.openqa.selenium.NoSuchElementException e3)
			   
			   {
			 
				   Wait<WebDriver> wait8 = new FluentWait<WebDriver>(driver)
						   .withTimeout(30, TimeUnit.SECONDS)
						   .pollingEvery(1, TimeUnit.SECONDS)
						   .ignoring(NoSuchElementException.class);
				   
				   WebElement element = wait8.until(new Function<WebDriver, WebElement>()
						   {
					         public WebElement apply(WebDriver driver)
					         {
					        	 return driver.findElement(By.linkText("Help"));
					         }
					   
						   });
				   
					
		        Actions builder = new Actions(driver); 
		        WebElement mainmenu = driver.findElement(By.linkText("Help"));
		        builder.moveToElement(mainmenu).build().perform();
		     
		        WebDriverWait wait1 = new WebDriverWait(driver,5000);
		        wait1.until(ExpectedConditions.visibilityOfElementLocated(By.linkText("Version")));
		        WebElement submenu=  driver.findElement(By.linkText("Version")); 
		        builder.moveToElement(submenu).click().build().perform();
		        Thread.sleep(3000);
		        
		        //Capture Application Version and write into excel
		        driver.switchTo().frame(driver.findElement(By.name("aeRadModalWindow")));
		        Thread.sleep(3000);
		        driver.findElement(By.className("textboxlabelleft1"));
		        String AppVersion = driver.findElement(By.id("lblApplicationVersion")).getText();
		        Cell c = s.getRow(row).getCell(4);
		        c.setCellValue(AppVersion);
		       
		        if(AppVersion.equals(s.getRow(row).getCell(3).getStringCellValue()))
				{
		        	s.getRow(row).createCell(5).setCellValue("Passed");		        					
				}
				else
				{
					s.getRow(row).createCell(5).setCellValue("Failed");
				}
		             
		        
		        
		        //Capture DB Version and write into excel
		        
		        String DBVersion = driver.findElement(By.id("lblDatabaseVersion")).getText();
		        c = s.getRow(row).getCell(7);
		        c.setCellValue(DBVersion);
		        
		        if(DBVersion.equals(s.getRow(row).getCell(6).getStringCellValue()))
				{
		        	s.getRow(row).createCell(8).setCellValue("Passed");	        					
				}
				else
				{
					s.getRow(row).createCell(8).setCellValue("Failed");
				}
		             
		     	       
		        
		        FileOutputStream fos = new FileOutputStream("C:/SCS/Ananth/AEHostedURLsValidation/HostedURLs.xlsx");
		        w.write(fos);
		        fos.close();
		        
		        driver.findElement(By.id("cmdOK")).click();
		        Thread.sleep(5000);
		        //driver.findElement(By.id("lnkLogout")).click(); 
		       		        
		        //Report execution
		        
		        Actions builder1 = new Actions(driver); 
		        WebElement mainmenu1 = driver.findElement(By.linkText("Reports"));
		        builder1.moveToElement(mainmenu1).build().perform();
		        
		        WebDriverWait wait2 = new WebDriverWait(driver,5000);
		        wait2.until(ExpectedConditions.visibilityOfElementLocated(By.linkText("Standard")));
		        WebElement submenu1=  driver.findElement(By.linkText("Standard")); 
		        builder1.moveToElement(submenu1).click().build().perform();
		        
		        WebDriverWait wait6 = new WebDriverWait(driver, 120);

			    wait6.until(new ExpectedCondition<Boolean>() {
			        public Boolean apply(WebDriver driver) {
			            return ((JavascriptExecutor) driver).executeScript(
			                "return document.readyState"
			            ).equals("complete");
			        }
			    });
		        
		        //Thread.sleep(10000);
		       
		        driver.findElement(By.xpath("//*[@id='ctl00ContentPlaceHolder1ReportTreeView_1']/img[1]")).click();
		        Thread.sleep(3000);
		        driver.findElement(By.xpath("//*[@id='ctl00ContentPlaceHolder1ReportTreeView_1_1']/span[3]")).click();
		        Thread.sleep(3000);
		        driver.findElement(By.id("ctl00_ContentPlaceHolder1_cmdRun")).click();
		        Thread.sleep(3000);
		        driver.findElement(By.id("ctl00_ContentPlaceHolder1_cmdRun")).click();
		        Thread.sleep(20000);
		        
		        String parentWindow = driver.getWindowHandle();
		        Set<String> handles =  driver.getWindowHandles();
		           for(String windowHandle  : handles)
		               {
		               if(!windowHandle.equals(parentWindow))
		                  {
		                  driver.switchTo().window(windowHandle);
		                 //Perform your operation here for new window
		                 Thread.sleep(5000); 
		                 	                 
		                 driver.close(); //closing child window
		                 driver.switchTo().window(parentWindow); //cntrl to parent window
		                 Thread.sleep(5000);
		                  }
		               }
		           
		            Actions builder2 = new Actions(driver); 
			        WebElement mainmenu2 = driver.findElement(By.linkText("Reports"));
			        builder1.moveToElement(mainmenu2).build().perform();
			     
			        WebDriverWait wait3 = new WebDriverWait(driver,5000);
			        wait3.until(ExpectedConditions.visibilityOfElementLocated(By.linkText("Advanced")));
			        WebElement submenu2=  driver.findElement(By.linkText("Advanced")); 
			        builder1.moveToElement(submenu2).click().build().perform();
			        Thread.sleep(5000);
			        
			        String parentWindow1 = driver.getWindowHandle();
			        Set<String> handles1 =  driver.getWindowHandles();
			           for(String windowHandle1  : handles1)
			               {
			               if(!windowHandle1.equals(parentWindow1))
			                  {
			                  driver.switchTo().window(windowHandle1);
			                 //Perform your operation here for new window
			                 
			                  WebDriverWait wait7 = new WebDriverWait(driver, 120);

			  			    wait7.until(new ExpectedCondition<Boolean>() {
			  			        public Boolean apply(WebDriver driver) {
			  			            return ((JavascriptExecutor) driver).executeScript(
			  			                "return document.readyState"
			  			            ).equals("complete");
			  			        }
			  			    });
			                  
			                  //Thread.sleep(5000);
			                 
			                 driver.findElement(By.xpath(".//*[@id='ctl00_createNew']/button")).click();
                             
                             
                            WebDriverWait wait4 = new WebDriverWait(driver,5000);
         			        wait4.until(ExpectedConditions.visibilityOfElementLocated(By.linkText("Report")));
         			        WebElement submenu3=  driver.findElement(By.linkText("Report")); 
         			        builder1.moveToElement(submenu3).click().build().perform();
         			        Thread.sleep(5000);
                             
                             
			                 
			                 Select DataSource_dd = new Select(driver.findElement(By.xpath(".//*[@id='ctl00_PlaceHolder_Adhocreportdesigner1_ctl01_jtc']/tbody/tr/td[3]/select")));
			                 DataSource_dd.selectByVisibleText("Account_View");
			                 Thread.sleep(5000);
			                 
			                 driver.findElement(By.id("continueBtn0")).click();
			                 Thread.sleep(5000);
			                 
			                 Select Field_dd = new Select(driver.findElement(By.xpath(".//*[@id='ctl00_PlaceHolder_Adhocreportdesigner1_ctl01_sc_Column']")));
				     	     Field_dd.selectByVisibleText("Account_Key");
				     	     Thread.sleep(5000);
				     	   
				     	  driver.findElement(By.xpath(".//*[@id='PreviewBtn0']")).click();
				     	  
				     	  Thread.sleep(20000); 
				     	  driver.close(); //closing child window
	                      driver.switchTo().window(parentWindow1); //cntrl to parent window
	                      Thread.sleep(5000);
			                 
			               } 
			        
		            	      		
		           
		      		}
			         driver.findElement(By.id("ctl00_HeaderControl1_lnkLogout")).click();
			}
			
			   }
			}
		}
}
		

		
       
		        
		        
			   
		
			

	


