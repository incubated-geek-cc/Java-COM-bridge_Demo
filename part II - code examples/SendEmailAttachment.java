import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import java.io.File;

public class SendEmailAttachment {
    public static void main(String[] args) {
		ActiveXComponent axOl = new ActiveXComponent("Outlook.Application");
        try {
            System.out.println("Version: " + axOl.getPropertyAsString("Version"));
            System.out.println("[Start] Sending Outlook Email");
            Dispatch olo = axOl.getObject();

            Dispatch mailItem = Dispatch
                    .invoke(olo, 
                            "CreateItem", 
                            Dispatch.Get, 
                            new Object[]{"0"}, 
                            new int[0])
                    .toDispatch();
            /* TO DO */
            /* Edit fields accordingly ---------------------*/
            String recipientEmailAddress="xxx@gmail.com";
            String attachmentFilepath="filename.png";

            Dispatch.put(mailItem, "To", recipientEmailAddress); 
            Dispatch.put(mailItem, "Subject", "JaCoB - Testing Send Email");
            Dispatch.put(mailItem, "Body", "Status: Success");
            Dispatch.put(mailItem, "ReadReceiptRequested", "false");

            /* (Optional) Include file attachment in email */
            Dispatch attachments = Dispatch
                    .get(mailItem, "Attachments")
                    .toDispatch();
            File object=new File(attachmentFilepath);
            Object obj=new Object();
            obj=object.getAbsolutePath();

            Dispatch.call(attachments, "Add", obj);
            Dispatch.call(mailItem, "Send");
            System.out.println("[...]" + "\n");
        } finally {
            System.out.println("[End] Sent Outlook Email" + "\n");
            axOl.invoke("Quit", new Variant[] {}); // (optional - to close app)
        }
	}
}