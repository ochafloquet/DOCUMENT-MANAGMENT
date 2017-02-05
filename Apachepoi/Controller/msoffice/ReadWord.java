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

import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
/**
 *
 * @author Oscar Chafloque
 */
public class ReadWord {
    
    public static void main(String[] args){
        
        try{
            FileInputStream fis= new FileInputStream("H:\\OFICIOTEMPLATE.docx");
            XWPFDocument docx = new XWPFDocument(fis);
            
            
            List<XWPFParagraph> paragraphList = docx.getParagraphs();
            int nump = paragraphList.size();
            System.out.println(nump); 
            
          
            
            for (int x = 0; x < paragraphList.size(); x++) { 
                
            	
            	
            }
            
            
            
            
            for(XWPFParagraph paragraph:paragraphList){
            	
            	String text = paragraph.getText();
            	
            	 XWPFRun rh = paragraph.createRun();
            	 
            	 
            	
            	if (text.contains("unidad")) 
                {
                    text = text.replace("unidad", "CICTE/W-6.a/02.00");
                    paragraph.removeRun(8);
                    rh.setText(text);

                }
            	 if (text.contains("receptor")) 
                 {
                     text = text.replace("receptor", "Gral Brig Jefe del Servicio de Material de Guerra del EjÃ©rcito");
                    
                 }

                 if (text.contains("asunto")) 
                 {
                     text = text.replace("asunto", "Sobre articulo de MG (Armamento) y apoyo de elemento tecnico.");

                 }
                 
                 if (text.contains("referencia")) 
                 {
                     text = text.replace("referencia", "Oficio N°289/CICTE del 01 julio de 2015.");

                 }

                 if (text.contains("cuerpo")) 
                 {
                     text = text.replace("cuerpo", "Tengo el honor de dirigirme a Ud., para manifestarle que en relación a la solicitud de prestamo de una (01) ametralladora BROWNING Cal .50 y la participaciÃ³n del elemento tÃ©cnico Tco 2da MAM Pacheco Tejada Henry, para las pruebas del vehÃ­culo blindado â€œOTORONGOâ€�, las cuales se han suspendido y serÃ¡n reprogramadas.\n" +
 "Asimismo, se informarÃ¡ de manera oportuna la fecha de realizaciÃ³n de las pruebas del vehÃ­culo blindado OTORONGO, para poder contar con artÃ­culo de MG (Armamento) y apoyo de elemento tÃ©cnico solicitado.\n" +
 "Hago propicia la oportunidad para expresarle a Ud. los sentimientos de mi especial consideraciÃ³n y estima personal.");

                 }
            	
            	System.out.println(text);
            	docx.write(new FileOutputStream("OFICIO.docx"));
            }
            
            
        }catch(FileNotFoundException e){
        e.printStackTrace();
        }catch(IOException e){
        e.printStackTrace();
        }
    
    }
    
}
