package msoffice;


import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;




public class prueba {

	public static void main(String[] args) throws InvalidFormatException, IOException {
		// TODO Auto-generated method stub
		XWPFDocument doc = new XWPFDocument(OPCPackage.open("OFICIOTEMPLATE.docx"));
		for (XWPFParagraph p : doc.getParagraphs()) {
		    List<XWPFRun> runs = p.getRuns();
		    if (runs != null) {
		        for (XWPFRun r : runs) {
		            String text = r.getText(0);
		            
		            if (text != null && text.contains("unidad")) {
		                text = text.replace("unidad", "CICTE/W-6.a/02.00");
		                r.setText(text, 0);
		            }
		            if (text != null && text.contains("receptor")) {
		                text = text.replace("receptor", "Gral Brig Jefe del Servicio de Material de Guerra del EjÃ©rcito");
		                r.setText(text, 0);
		            }
		            if (text != null && text.contains("asunto")) {
		                text = text.replace("asunto", "Sobre articulo de MG (Armamento) y apoyo de elemento tecnico.");
		                r.setText(text, 0);
		            }
		            if (text != null && text.contains("referencia")) {
		                text = text.replace("referencia", "Oficio N°289/CICTE del 01 julio de 2015.");
		                r.setText(text, 0);
		            }
		            if (text != null && text.contains("cuerpo")) {
		                text = text.replace("cuerpo", "Tengo el honor de dirigirme a Ud., para manifestarle que en relación a la solicitud de prestamo de una (01) ametralladora BROWNING Cal .50 y la participaciÃ³n del elemento tÃ©cnico Tco 2da MAM Pacheco Tejada Henry, para las pruebas del vehÃ­culo blindado â€œOTORONGOâ€�, las cuales se han suspendido y serÃ¡n reprogramadas.\n" +
		                		"Asimismo, se informarÃ¡ de manera oportuna la fecha de realizaciÃ³n de las pruebas del vehÃ­culo blindado OTORONGO, para poder contar con artÃ­culo de MG (Armamento) y apoyo de elemento tÃ©cnico solicitado.\n" +
		                		"Hago propicia la oportunidad para expresarle a Ud. los sentimientos de mi especial consideraciÃ³n y estima personal.");
		                r.setText(text, 0);
		            }
		        }
		    }
		}
		
		CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
		XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(doc, sectPr);
		
		//write header content
				CTP ctpHeader = CTP.Factory.newInstance();
			        CTR ctrHeader = ctpHeader.addNewR();
				CTText ctHeader = ctrHeader.addNewT();
				String headerText = "“AÑO DE LA DIVERSIFICACIÓN PRODUCTIVA Y DEL FORTALECIMIENTO DE LA EDUCACIÓN”";
				ctHeader.setStringValue(headerText);	
				XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, doc);
			        XWPFParagraph[] parsHeader = new XWPFParagraph[1];
			        parsHeader[0] = headerParagraph;
			        headerParagraph.setAlignment(ParagraphAlignment.CENTER);
			        policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);
			        
				//write footer content
				CTP ctpFooter = CTP.Factory.newInstance();
				CTR ctrFooter = ctpFooter.addNewR();
				CTText ctFooter = ctrFooter.addNewT();
				String footerText = "AÑO DE LA DIVERSIFICACIÓN PRODUCTIVA Y DEL FORTALECIMIENTO DE LA EDUCACIÓN";
				ctFooter.setStringValue(footerText);
				
				
				XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooter, doc);
			        XWPFParagraph[] parsFooter = new XWPFParagraph[1];
			        parsFooter[0] = footerParagraph;
			        footerParagraph.setAlignment(ParagraphAlignment.CENTER);
				policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, parsFooter);

		
		
		doc.write(new FileOutputStream("output.docx"));
		
		
	}

}
