/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package msoffice;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 *
 * @author Oscar Chafloque
 */
public class WriteinTemplate {
    
    public static void main(String[] args) throws FileNotFoundException, InvalidFormatException, IOException {
        
    	
    	
    	
      FileInputStream fis= new FileInputStream("H:\\OFICIOTEMPLATE.docx");
      XWPFDocument doc = new XWPFDocument(fis);
      
      for (XWPFParagraph p : doc.getParagraphs()) {
    	  
    	  
            for (XWPFRun r : p.getRuns()) {
            	//System.out.println(r);
                String text = r.getText(0);

                System.out.println(text);
                if (text.contains("unidad")) 
                {
                    text = text.replace("unidad", "CICTE/W-6.a/02.00");

                    r.setText(text, 0);
                    System.out.println(text);
                }

                if (text.contains("#receptor#")) 
                {
                    text = text.replace("#receptor#", "Gral Brig Jefe del Servicio de Material de Guerra del EjÃ©rcito");

                    r.setText(text, 0);
                    System.out.println(text);
                }

                if (text.contains("#asunto#")) 
                {
                    text = text.replace("#asunto#", "Sobre artÃ­culo de MG (Armamento) y apoyo de elemento tÃ©cnico.");

                    r.setText(text, 0);
                    System.out.println(text);
                }
                
                if (text.contains("#referencia#")) 
                {
                    text = text.replace("#referencia#", "Oficio NÂ°289/CICTE del 01 julio de 2015.");

                    r.setText(text, 0);
                    System.out.println(text);
                }

                if (text.contains("#cuerpo#")) 
                {
                    text = text.replace("#cuerpo#", "Tengo el honor de dirigirme a Ud., para manifestarle que en relaciÃ³n a la solicitud de prÃ©stamo de una (01) ametralladora BROWNING Cal .50 y la participaciÃ³n del elemento tÃ©cnico Tco 2da MAM Pacheco Tejada Henry, para las pruebas del vehÃ­culo blindado â€œOTORONGOâ€�, las cuales se han suspendido y serÃ¡n reprogramadas.\n" +
"Asimismo, se informarÃ¡ de manera oportuna la fecha de realizaciÃ³n de las pruebas del vehÃ­culo blindado â€œOTORONGOâ€�, para poder contar con artÃ­culo de MG (Armamento) y apoyo de elemento tÃ©cnico solicitado.\n" +
"Hago propicia la oportunidad para expresarle a Ud. los sentimientos de mi especial consideraciÃ³n y estima personal.");

                    r.setText(text, 0);
                    System.out.println(text);
                }

            }


         }
        doc.write(new FileOutputStream("output.docx"));


    
    }
    
}
