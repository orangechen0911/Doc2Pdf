
import java.io.File;

public class test {
    public static void main(String[] args) {
        //文件名
        String filename = "test.doc";
//        String filename = "test.docx";
        //文件地址
        String filePath = "C:";

        if (filename.endsWith(".doc")) {
            try {
                Doc2Pdf.docToPdf(filePath + File.separator + filename);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        if (filename.endsWith(".docx")) {
            try {
                Doc2Pdf.docxToPdf(filePath + File.separator + filename);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
}