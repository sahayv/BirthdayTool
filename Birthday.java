package birthdaytoolgss;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;
import java.util.Random;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Birthday {
	private static final String FILE_NAME = "C:/Users/sahayv/Desktop/BirthdayTool/DOBGSS.xlsx";
	  static String	 UserNamebirthday= " ";
  public static void main( String[] args )
  {
      System.out.println( "Checking birthday " );
      System.out.println("++++++++++++++");
      
      try {

          FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
          Workbook workbook = new XSSFWorkbook(excelFile);
          Sheet datatypeSheet = workbook.getSheetAt(0);
          Iterator<Row> iterator = datatypeSheet.iterator();

          while (iterator.hasNext()) {

              Row currentRow = iterator.next();
              Iterator<Cell> cellIterator = currentRow.iterator();

              while (cellIterator.hasNext()) {

                  Cell currentCell = cellIterator.next();
                  //getCellTypeEnum shown as deprecated for version 3.15
                  //getCellTypeEnum ill be renamed to getCellType starting from version 4.0
               
                  if (currentCell.getCellType() == CellType.STRING) {
                  //    System.out.print(currentCell.getStringCellValue() + "--");
                  	
                 	UserNamebirthday=currentCell.getStringCellValue();
                  } else if (currentCell.getCellType() == CellType.NUMERIC) {
                   //   System.out.print(currentCell.getDateCellValue() + "--");
                      Date birthday= currentCell.getDateCellValue();
                     // System.out.println(birthday);
                                                     
                      SimpleDateFormat sdf=new SimpleDateFormat("MM dd");
                      sdf.format(birthday);
                      
                      
                 String birthdayDate=sdf.format(birthday);
                    
                    Calendar cal=new GregorianCalendar();
                    String _todayDate= sdf.format(cal.getTime());
                    
                    if(birthdayDate.equals(_todayDate)){
                  	  System.out.println(_todayDate);
                  	  Cell currentCellUser = cellIterator.next();
                  	  System.out.println(currentCellUser.getStringCellValue()+"@vmware.com");
                  	  String userbirthday=currentCellUser.getStringCellValue()+"@vmware.com";
                  	  sendEmail(userbirthday,UserNamebirthday);
                    }
                      
              //   System.out.println(cal.get(field));
                 
                      
                  }

              }
          //    System.out.println();

          }
      } catch (FileNotFoundException e) {
          e.printStackTrace();
      } catch (IOException e) {
          e.printStackTrace();
      }

  }
  
  //random content 
  
  public static String givencontent() {
      
  	String randomElement= " ";
  	 
  	Random rand = new Random();
    
		List<String> givenList = new ArrayList<String>();
		givenList.add(0, "May all of your birthday wishes come true.");
		givenList.add(1,"Have a great time with all that you love.");
		givenList.add(2,"Be fancy and dandy this year");
		givenList.add(3,"Have a great time with all that you love.");
		givenList.add(4,"Do a little dance today.");
		givenList.add(5,"May your birthday be the start of a year filled with good luck, good health and much happiness.");
		givenList.add(6,"May all life's blessings be yours, on your birthday and always.");
		givenList.add(7,"The warmest wishes to a great member of our GSS Team. May your special day be full of happiness, fun and cheer!");
		
   
      int numberOfElements = 7;
   
      for (int i = 0; i < numberOfElements; i++) {
          int randomIndex = rand.nextInt(givenList.size());
           randomElement = givenList.get(randomIndex);
         
     //     givenList.remove(randomIndex);
      }
      return randomElement;
  }
  //Birthday poem

public static String givenPoem() {
      
  	String randomElement= " ";
  	 
  	Random rand = new Random();
    
		List<String> givenList = new ArrayList<String>();
		givenList.add(0, "Your age is just a number,<br>   " 
				+ "It’s really not all that big ,  <br>    "
				+ "You have seen a few decades, <br>   "
				+ "Since you were just a kid.  <br>  "
				+ "But you don't look your age,  <br>  "
				+ "You beat the middle age spread, <br>   "
				+ "Your face glows like 1,000 stars, <br>   "
				+ "And you can still turn heads. <br>   "
				+ "Best wishes and happy birthday! ");



		givenList.add(1,"On your birthday we wish you much pleasure and joy, <br>"
				+"We hope all of your wishes come true.<br>"
				+"May each hour and minute be filled with delight,<br>"
				+"And your birthday be perfect for you!") ;

		givenList.add(2,"This heartfelt wish is just for you<br>"
				+ "Today is your special day<br>"
				+ "May all the dreams you do pursue<br>"
				+ "Be realized in every way");


		givenList.add(3,"May this birthday be your best birthday ever, <br>  "
				+ "full of light and laughter,<br> "
				+ "a fireworks explosion of joy.<br> "
				+ "May this birthday live in your memory,<br> "
				+ "forever creating happiness and,<br> "
				+ "satisfaction whenever you remember it.<br> "
				+ "Happy, happy birthday!");

		givenList.add(4,"Another birthday, another year<br> "
				+ "May you have problems that disappear<br> "
				+ "May you have health<br> "
				+ "And a bit of wealth<br> "
				+ "May you share<br> "
				+ "With those who care<br> "
				+ "May the coming year <br> "
				+ "Be one of good cheer.");


		givenList.add(5,"Here is a birthday wish for you<br> "
				+ "One that is loving, happy and true<br> "
				+ "A wish for happiness and best things<br> "
				+ "The coming year to you will bring.<br> ");

		
		givenList.add(6,"H - is for the Happiest of all days  <br>  "
				+ "A - is for All the wishes and praise<br> "
				+ "P - is for the Presents you'll open with delight<br> "
				+ "P - is for the Party that will last into the night<br> "
				+ "Y - is for the Year leading up to your day<br> "
				+ "<br> "
				+ "<br> "
				+ "B - is for the Balloons a celebration they'll say<br> "
				+ "I - is for the Ice cream to have with your cake<br> "
				+ "R - is for the Ribbons and decorations you'll make<br> "
				+ "T - is for the Theme you'll decided to throw<br> "
				+ "H - is for the Hats made with confetti and a bow<br> "
				+ "<br> "
				+ "<br> "
				+ "D - is for the Day you know will be fun<br> "
				+ "A - is for Another great year that is done<br> "
				+ "Y - is for Your special day.<br> "
				+ "<br> "
				+ "<br> "
				+ "Happy Birthday! Happy Birthday! Hip-hip hooray!");


		givenList.add(7,"Today is your birthday,Today is your day,<br> "
				+ "To be happy in each and every way,<br> "
				+ "Be happy keep smiling in every thing you do,<br> "
				+ "Cause now I wish HAPPY BIRTHDAY to you!!!");
   
      int numberOfElements = 7;
   
      for (int i = 0; i < numberOfElements; i++) {
          int randomIndex = rand.nextInt(givenList.size());
           randomElement = givenList.get(randomIndex);
         
     //     givenList.remove(randomIndex);
      }
      return randomElement;
  }
  
  // Sending email
  public static void sendEmail(String user,String userName){
  	

  	   try {
             String to=user ;
             String from="Sahayv@vmware.com";
             String bcc="Sahayv@vmware.com";

             Properties props = new Properties();
//             props.put("mail.smtp.socketFactory.port", "587");
//             props.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
//             props.put("mail.smtp.socketFactory.fallback", "true");
             props.put("mail.smtp.host", "smtp.office365.com");
             props.put("mail.smtp.port", "587");
             props.put("mail.smtp.starttls.enable","true");
             props.put("mail.smtp.auth", "true");

             Session session = Session .getDefaultInstance(props,new javax.mail.Authenticator() { protected javax.mail.PasswordAuthentication getPasswordAuthentication() {return new javax.mail.PasswordAuthentication("sahayv@vmware.com","Om1namah*0001");
                         }
                     });

//             Session emailSession = Session.getDefaultInstance(props, null);

           //  String msgBody = "Sending email using JavaMail API...";
           String[] userFirtsName= userName.split(" ");
             
             String  message = "<body bgcolor=Azure></body>    <H1 align=center > <font color=DarkBlue     size=7  align=center ><b>  "+"Happy Birthday "+ userFirtsName[0] +" !!" +"</font> </b><br>";
             message += "<H4><font color=DarkRed  size=6><b>" +  givencontent() +"</font> </b><br> ";
             message += "<div  style=background-color:powderblue   align=center  ><i>"+givenPoem()+"</i></div>";
             
          message +="<img src=\"C:/Users/sahayv/Desktop/BirthdayTool/Pic/image.jpeg\" height=600 width=1200> ";
             
             
             Message msg = new MimeMessage(session);
             msg.setFrom(new InternetAddress(from, "NoReply"));
             msg.addRecipient(Message.RecipientType.TO,
                     new InternetAddress(to, "Mr. Recipient"));
             msg.addRecipient(Message.RecipientType.BCC,
                     new InternetAddress(bcc, "Mr. Bcced"));
             msg.setSubject("Happy Birthday "+userFirtsName[0]+"!!");
            
             msg.setContent(message ,"text/html"); 
           //  msg.setText(msgBody);
             Transport.send(msg);
             System.out.println("Email sent successfully...");
          //   logger.error("Email sent successfully...");
         } catch (AddressException e) {
           //  logger.error(e.getMessage());
         } catch (MessagingException e) {
            // logger.error(e.getMessage());
         } catch (UnsupportedEncodingException e) {
            // logger.error(e.getMessage());
         }
  	
  	
  }
  }
