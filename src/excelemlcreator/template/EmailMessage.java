package excelemlcreator.template;

import org.apache.poi.ss.usermodel.Row;

import javax.mail.*;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Properties;

public abstract class EmailMessage {
    private MimeMessage message;
    public EmailMessage(){ }
    public MimeMessage createNewEmptyMessage(){
        return new MimeMessage((Session) null);
    }
    public void setMessage(MimeMessage message){
        this.message = message;
    };
    public MimeMessage getMessage (){
        return message;
    }
    public boolean addRecipientTO(String email){
        return addRecipient(email, Message.RecipientType.TO);
    }
    public boolean addRecipientCC(String email){
        return addRecipient(email, Message.RecipientType.CC);
    }
    private boolean addRecipient(String email, Message.RecipientType type){
        try{
            message.addRecipient(type, new InternetAddress(email));
            return true;
        }catch(AddressException e){
            System.out.println("invalid email address for: "+email);

        }catch (MessagingException e){
            System.out.println("Not able to add recipient using email: "+email);
        }
        return false;
    }
    public boolean setFrom(String fromEmailAddress){
        try{
            message.setFrom(fromEmailAddress);
            return true;
        }catch(MessagingException e){
            System.out.println("Not able to set subject using: "+fromEmailAddress);
        }
        return false;
    }
    public boolean setSubject(String subject){
        try{
            message.setSubject(subject);
            return true;
        }catch(MessagingException e){
            System.out.println("Not able to set subject using: "+subject);
        }
        return false;
    }
    public boolean setBodyContent(String bodyContent){
        try{
            message.setContent(bodyContent, "text/html");
            return true;
        }catch(MessagingException e){
            System.out.println("Not able to set body content using: "+bodyContent);
        }
        return false;
    }
    public abstract ArrayList<MimeMessage> createMimeMessages(String excelSrc);
}
