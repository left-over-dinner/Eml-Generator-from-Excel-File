package excelemlcreator.template;

import org.apache.poi.ss.usermodel.Row;

import javax.mail.*;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import java.util.Iterator;
import java.util.Properties;

public abstract class EmailMessage {
    private MimeMessage message;
    public EmailMessage(){
        message = new MimeMessage((Session) null);
    }
    public MimeMessage getMessage (){
        return message;
    }
    public void resetMessageAttr(){
        try{
            message.setSubject(null);
            message.setContent((Multipart) null);
            message.setRecipient( Message.RecipientType.TO, null);
            message.setRecipient( Message.RecipientType.CC, null);
        }catch(MessagingException e){
            System.out.println("error in resetting the eml");
        }
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
    public boolean setSubject(String subject){
        try{
            message.setSubject(subject);
            return true;
        }catch(MessagingException e){
            System.out.println("Not able to set subject using: "+subject);
        }
        return false;
    }
    public boolean setBodyContent(EmailBodyContent bodyContent){
        try{
            message.setContent(bodyContent.createBody(), "text/html");
            return true;
        }catch(MessagingException e){
            System.out.println("Not able to set body content using: "+bodyContent);
        }
        return false;
    }
    public abstract MimeMessage createMimeMessage(Iterator<Row> dataRows);
}
