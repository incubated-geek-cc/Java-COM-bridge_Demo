import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import java.util.Date;

public class RecurseEmailFolders {
    public static void main(String[] args) {
    ActiveXComponent axOl = new ActiveXComponent("Outlook.Application");
    try {
      System.out.println("Version: " + axOl.getPropertyAsString("Version"));
      System.out.println("[Start] Browsing Outlook Email folders");
      Dispatch olo = axOl.getObject();
      Dispatch namespace = Dispatch
                              .call(olo, "GetNamespace", "MAPI")
                              .toDispatch();
      browseFolders(0, namespace);
    } finally {
        System.out.println("[End] Browse Outlook Email folders" + "\n");
    }
  }
  private static String pad(int i) {
    StringBuffer sb = new StringBuffer();
    while (sb.length() < i) {
        sb.append(' ');
    }
    return sb.toString();
  }
  private static void browseFolders(int iIndent, Dispatch namespace) {
    Dispatch folders = Dispatch.get(namespace, "Folders").toDispatch();
    int folderCount = folders.call(folders, "Count").toInt(); 

    Dispatch folder=null;
    String name = null;

    int folderCounter=0;
    while (folderCounter < folderCount) {
        if (folderCounter == 0) {
            folder = Dispatch.get(folders, "GetFirst").toDispatch();
        } else {
            folder = Dispatch.get(folders, "GetNext").toDispatch();
        }
        name = folder.call(folder, "Name").toString();
        System.out.println(pad(iIndent) + name);

        Dispatch items = Dispatch.get(folder, "Items").toDispatch(); 
        int count = Dispatch.call(items, "Count").toInt();
        System.out.println("Total item count: " + count + "\n");

        Dispatch item = null;
        String msg_id = null;
        String msg_subject = null;
        String msg_sender_name = null;
        String msg_sender_email = null;
        Date msg_received_date = null;

        System.out.println("[Start] Reading items in " + name);
        int itemCounter = 0;
        while (itemCounter < count) {
          if (itemCounter == 0) {
              item = Dispatch.get(items, "GetFirst").toDispatch();
          } else {
              item = Dispatch.get(items, "GetNext").toDispatch();
          }
          try {
              msg_id = item.call(item, "EntryID").toString();
              msg_subject = item.call(item, "Subject").toString();
          } catch (Exception ex) {
              System.out.println(ex);
          } finally {
              itemCounter++;
          }
        }
        System.out.println("[End] Read items in " + name);
        folderCounter++;
        /* Browses recursively */
        browseFolders(iIndent + 3, folder);
    }
  }
}
