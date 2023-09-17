package com;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class DownloadEmailAttachments {
    public static void main(String[] args) {
		ActiveXComponent axO = new ActiveXComponent("Outlook.Application");
        System.out.println("Version: " + axO.getPropertyAsString("Version") + "\n");
        System.out.println("[Start] Downloading Outlook Email Inbox attachments" + "\n");
        try {
            Dispatch olo = axO.getObject();
            Dispatch namespace = Dispatch.call(olo, "GetNamespace", "MAPI").toDispatch();
            Dispatch inbox = Dispatch.call((Dispatch) namespace, 
                                           "GetDefaultFolder", 
                                           // '6' refers to inbox
                                           new Integer(6)).toDispatch(); 
            Dispatch items = Dispatch.get(inbox, "Items").toDispatch();

            int count = Dispatch.call(items, "Count").toInt();
            int itemCounter=0;
            Dispatch item=null;
            while (itemCounter < count) {
                if (itemCounter == 0) {
                    item = Dispatch.get(items, "GetFirst").toDispatch();
                } else {
                    item = Dispatch.get(items, "GetNext").toDispatch();
                }
                try {
                    Dispatch emailAttachments = Dispatch.get(item, "Attachments")
                                                        .toDispatch();
                    int attachmentCount = Dispatch.call(emailAttachments,"Count")
                                                    .toInt();
                    System.out.println("Attachment count: "+ attachmentCount + "\n");
                    int attachmentCounter=1;
                    while(attachmentCounter<=attachmentCount) {
                        Dispatch attachmentFile = Dispatch.call(emailAttachments, 
                                                                "Item", 
                                                                attachmentCounter)
                                                            .toDispatch();
                        String displayFileName=Dispatch
                                                    .call(attachmentFile, 
                                                          "DisplayName")
                                                    .toString();
                        Dispatch attachmentItem=Dispatch
                                                    .call(emailAttachments, 
                                                            "Item", 
                                                            attachmentCounter).toDispatch();
                        String dirPath=System.getProperty("user.dir")+"\\output";

                        Path folder = Paths.get(dirPath);
                        if (!Files.exists(folder)) {
                            Files.createDirectories(Paths.get(dirPath));
                        }
                        String saveLocation=dirPath+"\\"+displayFileName;
                        System.out.println("File attachment saved at: "+ saveLocation);
                        Dispatch.call(attachmentItem, "SaveAsFile", saveLocation);
                        attachmentCounter++;
                    }
                } catch (Exception ex) {
                    System.out.println(ex);
                } finally {
                    itemCounter++;
                }
            }
        } finally {
            System.out.println("\n" + "[End] Download Outlook Email Inbox attachments" + "\n");
        }
	}
}
