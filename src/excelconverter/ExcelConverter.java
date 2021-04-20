/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelconverter;

import com.jcraft.jsch.ChannelSftp;
import com.jcraft.jsch.JSch;
import com.jcraft.jsch.JSchException;
import com.jcraft.jsch.Session;
import com.monitorjbl.xlsx.StreamingReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author HOME
 */
public class ExcelConverter {

    /**
     * @param args the command line arguments
     */

    public static ChannelSftp setupJsch(String host,String username,String password){
       try{ 
        JSch jsch = new JSch();
        Session jschSession = jsch.getSession(username, host);
        java.util.Properties config = new java.util.Properties();
        config.put("StrictHostKeyChecking", "no");
        jschSession.setConfig(config);
        jschSession.setPassword(password);
        jschSession.connect();
        return (ChannelSftp) jschSession.openChannel("sftp");
       }catch(JSchException e){
           return null;
       }
        
    }
    public static Boolean downloadExcel(ChannelSftp ftp,String sourcePath,String excelFile){
        // TODO download excel file 
        try{
          ftp.connect();
          ftp.get(sourcePath+excelFile,excelFile);
          ftp.rm(sourcePath+excelFile);
        }catch(Exception e){
           e.printStackTrace();
           System.out.println("Error while downloading");
           return false;
        }
        System.out.println("Download Complete");
        return true;
    }
    public static Boolean convertCSV(String fileName,String[] list){
        String filePath = fileName;
        try{
        int rowCacheSize = 100;
        InputStream is = new FileInputStream(filePath);
            Workbook workbook = StreamingReader.builder()
                    .rowCacheSize(rowCacheSize) // number of rows to keep in memory (defaults to 10)
                    .bufferSize(4096) // buffer size to use when reading InputStream to file (defaults to 1024)
                    .open(is);
               int sheetIndex = 0;
               for (Sheet sheet : workbook) {
                String name = sheet.getSheetName();
                if(sheetIndex< 2) list[sheetIndex] = name + ".csv";
                File file = new File(name+".csv");
                if(file.exists()) file.delete();
                OutputStream os = new FileOutputStream(name+".csv",true);
                String cacheString= "";
                int i=0,j=0;
                for (Row r : sheet) {
                    i++;
                    j = 0;
                    for (Cell c : r) {
                        String value = c.getStringCellValue();
                        if(value.contains(",")) value = "\""+value+"\"";
                        if(j == 0)
                            cacheString += value;
                        else
                            cacheString += "," + value;  
                        j++;
                    }
                    cacheString += "\n";
                    if(i == rowCacheSize){
                        byte[] b = cacheString.getBytes();
                        os.write(b);
                        i = 0;
                        cacheString = "";
                    }
                    
                }
                if(i > 0){
                    byte[] b = cacheString.getBytes();
                    os.write(b);                    
                }
                os.close();
                sheetIndex ++;
            }
        }
        catch(Exception e){
              e.printStackTrace();
        }
        return true;
    }
    public static Boolean uploadCSV(ChannelSftp ftp,String remotePath,String[] list,int count){
        // TODO upload csv files
        try {
            //ftp.connect();
        	for(int i = 0; i < count ; i++) {
            ftp.put(list[i],remotePath+list[i]);
            
        	}
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("Error while uploading");
            return false;
        }
        System.out.println("Upload Complete");
        return true;
    }
    public static void main(String[] args) {
        System.out.println("start program");
        if (args.length < 7) {
            System.out.println("Invalid password");
            return;
        }
        String host = args[0];
        int port = Integer.parseInt(args[1]);
        String username = args[2];
        String password = args[3];
        String sourcePath = args[4];
        String destPath = args[5];
        String excelFile = args[6];
        System.out.println("Host: " + host);
        System.out.println("Port: " + port);
        System.out.println("Username: " + username);
        System.out.println("Password: " + password);
        System.out.println("sourcePath: " + sourcePath);
        System.out.println("destPath: " + destPath);
        System.out.println("ExcelFile: " + excelFile);
       
        
        ChannelSftp ftp = setupJsch(host,username,password);
        if(ftp == null){
            System.out.println("not connected to sftp server");
            return;
        }
        /// downloaded excel file will save on current directory 
        if(!downloadExcel(ftp,sourcePath,excelFile)){
            System.out.print("Can't download excel file");
            return;
        }
        String[] list= new String[2];
        if(!convertCSV(excelFile,list)){
            System.out.print("Can't convert excel file");
            return;
        }
        System.out.println(list[0]);
        System.out.println(list[1]);
        if (!uploadCSV(ftp,destPath,list,2)) {
            System.out.print("Can't upload csv files to destination path");
            return;
        }
        if (!uploadCSV(ftp,sourcePath,list,1)) {
            System.out.print("Can't upload csv files to source path");
            return;
        }
        return;
        
    }
    
}
