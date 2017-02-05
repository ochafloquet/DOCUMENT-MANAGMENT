package msoffice;


import java.io.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

 
public class imagen
{
    public static void main(String[] args) throws IOException, InvalidFormatException
    {
    	XWPFDocument docx = new XWPFDocument();
    	XWPFParagraph par = docx.createParagraph();  
    	XWPFRun run = par.createRun();
    	run.setText("Hello, World. This is my first java generated docx-file. Have fun.");
    	run.setFontSize(13);

    	InputStream pic = new FileInputStream("firma jefe.jpg");
    	byte [] picbytes = IOUtils.toByteArray(pic);
    	docx.addPictureData(picbytes, Document.PICTURE_TYPE_JPEG);
    	docx.write(new FileOutputStream("output1.docx"));
    }
}


