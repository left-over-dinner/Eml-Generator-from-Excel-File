package excelemlcreator;

import javax.mail.MessagingException;
import javax.mail.internet.MimeMessage;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class EmlAdapter {
    private final String emlDateFormat = "yyyy-MM-dd_HH-mm-ss-SSS";
    private final String outputFolderDefaultName = "EmlMessages";
    private SimpleDateFormat dateFormatter;
    private String folderPath;
    private int writtenFileCount;
    public EmlAdapter(){
        dateFormatter = new SimpleDateFormat(emlDateFormat);
        writtenFileCount=0;
    }
    public void setOutputFolderPath(String folderPath){
        this.folderPath = folderPath;
    }
    public void createEmlMessages(MimeMessage[] mimeMessages){
        System.out.println("Creating eml file(s)");
        String outputFolderNameFinal = createDirectory();
        for(int i = 0; i<mimeMessages.length;i++){
            readMimeMessage(i,mimeMessages[i],outputFolderNameFinal);
        }
        System.out.println(writtenFileCount+" out of "+mimeMessages.length+" successful eml file(s) created");
        System.out.println("Eml file(s) location: "+outputFolderNameFinal);

    }
    private void readMimeMessage(int index, MimeMessage mimeMessage, String outputFolderNameFinal){
        String emlFileName = createEmlFileName(index,mimeMessage);
        if(emlFileName == null){ System.out.println("Unsuccessful eml creation for current mimeMessage at index "+index); return;}
        String outputPathFile=outputFolderNameFinal+File.separator+emlFileName+".eml";
        writeMimeMessageToFile(outputPathFile, mimeMessage);
        writtenFileCount++;
    }
    private String createDirectory(){
        boolean endingWithFileSeparator = (folderPath.charAt(folderPath.length()-1))==File.separatorChar;
        String fileSeparator = endingWithFileSeparator?"":File.separator;
        String pathToDirectory = folderPath+fileSeparator+outputFolderDefaultName+"_"+getCurrentDateAndTime();
        new File(pathToDirectory).mkdirs();
        return pathToDirectory;
    }
    private void writeMimeMessageToFile(String outputPath, MimeMessage mimeMessage){
        try{
            mimeMessage.writeTo(new FileOutputStream(outputPath));
        }catch(FileNotFoundException e){
            System.out.println("Unable to create file with name "+ outputPath + " file error");
        }catch(MessagingException e){
            System.out.println("Unable to write file with name "+ outputPath+ " mimeMessage error");
        }catch(IOException e){
            System.out.println("Unable to write file with name "+ outputPath+ " IO error");
        }

    }
    private String createEmlFileName(int index, MimeMessage mimeMessage){
        try{
            return (index+1)+" "+mimeMessage.getSubject()+" "+getCurrentDateAndTime();
        }catch (MessagingException e){
            return null;
        }

    }
    private String getCurrentDateAndTime(){
        return dateFormatter.format(new Date());
    }
}
