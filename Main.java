import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import java.util.Date;
import org.json.JSONArray;
import org.json.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

public class Main {
    static {
        System.load(System.getProperty("user.dir")+"\\jacob-1.18-x64.dll");
    }
    public static void main(String[] args) {
        // Step 0
        ActiveXComponent axOutlook = new ActiveXComponent("Outlook.Application");
        Dispatch olo = axOutlook.getObject();
        String version=axOutlook.getProperty("Version").getString();          
        System.out.println("Outlook Version: " + version);
        
        // Step 1
        System.out.println("---");
        System.out.println("[Start] Scanning Outlook Inbox");
        System.out.println("---");
        Dispatch namespace = Dispatch.call(olo, "GetNamespace", "MAPI").toDispatch();
        // '6' refers to inbox
        Dispatch inbox = Dispatch.call((Dispatch) namespace,"GetDefaultFolder", new Integer(6)).toDispatch();
        Dispatch items = Dispatch.get(inbox, "Items").toDispatch(); // .get(__, Method).toDispatch();
        int count = Dispatch.call(items, "Count").toInt(); // .call(__, Property).to<DataType>();
        System.out.println("Total Items in Inbox: " + count);
        System.out.println("---");
        
        // Step 2
        int itemCounter=0;
        Dispatch item;
        String msg_id = null;
        String msg_subject = null;
        String msg_sender_name = null;
        String msg_sender_email = null;
        Date msg_received_date = null;
        
        JSONArray arr = new JSONArray();
        while(itemCounter<count) {
            if(itemCounter==0) {
                item = Dispatch.get(items, "GetFirst").toDispatch();
            } else {
                item = Dispatch.get(items, "GetNext").toDispatch();
            }
            msg_subject = item.call(item,"Subject").toString();
            msg_id = item.call(item,"EntryID").toString();
            msg_sender_name = item.call(item,"SenderName").toString();
            msg_sender_email = item.call(item,"SenderEmailAddress").toString();
            msg_received_date = item.call(item,"ReceivedTime").toJavaDate();
            
            System.out.println("Email Details:");
            System.out.println("---");
            System.out.println("Item ID: "+msg_id+" ["+msg_received_date+"]");
            System.out.println("From: "+msg_sender_name + " <" + msg_sender_email + ">");
            System.out.println("Subject: "+msg_subject);
            System.out.println("---");
            
            // Step 3
            if(msg_subject.contains("Case Notification") && msg_subject.contains(("CaseID"))) {
                JSONObject obj = new JSONObject();
                String msg_html_body = item.call(item,"HTMLBody").toString();
                Document doc = Jsoup.parse(msg_html_body);
                Elements tableRows=doc.select("table tr");
                tableRows.forEach(ele -> {
                    Elements tds=ele.getElementsByTag("td");
                    int cols = tds.size();
                    if(cols==2) {
                        String headerText=tds.get(0).text();
                        String textContent=tds.get(1).text();
                        obj.put(headerText, textContent);
                    }
                });
                arr.put(obj);
            }
            itemCounter++;
        }
        System.out.println(arr.toString(2));
        System.out.println("[End] Reading Outlook Inbox Items");
        System.out.println("---");
    }
}
