package net.delpilar.textgenerator;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.awt.Desktop;

import org.apache.poi.ss.usermodel.*;


public class App 
{
    public static void main( String[] args ) throws IOException
    {
        Workbook workbook = WorkbookFactory.create(new File("/Users/jordandelpilar/Desktop/farts.xlsx"));
        Sheet sheet = workbook.getSheetAt(0);
        LocalDateTime today_raw = LocalDateTime.now();
        DateTimeFormatter date_format = DateTimeFormatter.ofPattern("MM-DD-YYYY");
        String today = today_raw.format(date_format);
        File file = new File(today + ".txt");
        FileWriter out_file = new FileWriter(file);

        String msg = 
"""
Artificial Grass Delivery Confirmation- Your order, %s, has been dispatched and will be delivered %s between %s - %s at %s.
To prepare for your delivery please make sure nothing is blocking the delivery location selected. 
You will receive another text notification 30 minutes prior to arrival. If there is a gate or entry approval, please provide and confirm.
""";

        for(Row row : sheet) {
            String value = "";

            try {
                value = row.getCell(1).getStringCellValue();
            }
            catch(Exception e) {
                System.out.printf("ERROR: %s", e);
                continue;
            }

            if(value.matches("^[0-9]*$") && value != "") {
                String workorder_num = value;
                String appointment_num = row.getCell(2).getStringCellValue();
                String customer_name = row.getCell(3).getStringCellValue();
                String address = row.getCell(5).getStringCellValue();

                String phone_num = row.getCell(8).getStringCellValue();
                phone_num = phone_num.replaceAll("[^0-9]", "");
                if(phone_num.length() > 0){
                    if(phone_num.startsWith("1")) {
                        phone_num = "+" + phone_num;
                    }
                    else {
                        phone_num = "+1" + phone_num;
                    }
                }

                if(phone_num.length() > 12) {
                    phone_num = phone_num.substring(0, 12);
                }
                
                String start_time_String = "";
                String end_time_String = "";
                int start_time = 0;
                int end_time = 0;
                boolean time_error = false;
                

                try {
                    start_time_String= row.getCell(7).getStringCellValue();
                }
                catch(IllegalStateException ignore) {
                    
                }
                
                try {
                    DateTimeFormatter datetime_format = DateTimeFormatter.ofPattern("M/d/y H:m a");
                    start_time = LocalDateTime.parse(start_time_String, datetime_format).getHour();
                }
                catch(DateTimeParseException ignore) {
                    try {
                    start_time = LocalDateTime.parse(start_time_String, DateTimeFormatter.ISO_DATE_TIME).getHour();
                    }
                    catch(DateTimeParseException e) {
                        start_time = 0;
                        time_error = true;
                    }
                }

                if(address.contains("FL")) {
                    start_time += 3;
                }

                if(start_time <= 3) {
                    start_time = 4;
                }

                end_time = start_time + 2;

                if(start_time > 12){
                    start_time_String = Integer.toString(start_time - 12) + " PM"; 
                }
                else if(start_time == 12) {
                    start_time_String = Integer.toString(start_time) + " PM"; 
                }
                else {
                    start_time_String = Integer.toString(start_time) + " AM";
                }

                if(end_time > 12){
                    end_time_String = Integer.toString(end_time - 12) + " PM"; 
                }
                else if(end_time == 12) {
                    end_time_String = Integer.toString(end_time) + " PM"; 
                }
                else {
                    end_time_String = Integer.toString(end_time) + " AM";
                }

                String delivery_day = "";

                if(today_raw.getDayOfWeek().getValue() == 5) {
                    delivery_day = "monday";
                }
                else {
                    delivery_day = "tomorrow";
                }

                msg = String.format(
                    msg, 
                    workorder_num, 
                    delivery_day, 
                    start_time_String, 
                    end_time_String, 
                    address
                    );

                if(time_error){
                    out_file.write("*** Could Not Parse Start Time For The Followng Record *** \n");
                }
                out_file.write(
                    String.format(
                        "WO #: %s\tAppointment #: %s\nCustomer Name: %s\nAddress: %s\nPhone Number: %s\nStart Time: %s\tEnd Time: %s\n\n%s==========================\n\n",
                        workorder_num,
                        appointment_num,
                        customer_name,
                        address,
                        phone_num,
                        start_time_String,
                        end_time_String,
                        msg
                        )
                    );
            } 
            else {
                continue;
            }            
        }

        out_file.close();
        Desktop.getDesktop().open(file);

    }
}
