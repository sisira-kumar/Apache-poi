package org.ssd;

import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabStop;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabs;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabJc;

public class Test {
	/*public static void main(String args[]) throws IOException
	{
		XWPFDocument document = new XWPFDocument();
		XWPFParagraph p1 = document.createParagraph();
		FileOutputStream out = new FileOutputStream(new File("F:\\ssd\\text1.docx"));
		p1.setFirstLineIndent(400);
		p1.setAlignment(ParagraphAlignment.CENTER);
		XWPFParagraph p2 = document.createParagraph();
		
		p2.setFirstLineIndent(800);
		p2.setAlignment(ParagraphAlignment.LEFT);
		
		XWPFRun r1 = p1.createRun();
		
		String t1 = "APPLICATION"+"\t\t\t"+"FORM";
		r1.setText(t1);
		r1=p2.createRun();
		r1.addCarriageReturn();
		r1.addCarriageReturn();
		r1.addCarriageReturn();
		r1.addCarriageReturn();
		r1.addCarriageReturn();
		r1.addCarriageReturn();
		r1.setFontFamily("Times New Roman");
		String t2="\n\nDesi ........";
		r1.setText(t2);
		document.write(out);
		out.close();
		document.close();
	}
*/
	public static void main(String[] args) {
        try {
        	@SuppressWarnings("resource")
			XWPFDocument document = new XWPFDocument();
        	XWPFParagraph p1 = document.createParagraph();
        	
        	p1.setFirstLineIndent(400);
        	p1.setAlignment(ParagraphAlignment.CENTER);
        	XWPFParagraph p2 = document.createParagraph();

        	p2.setFirstLineIndent(800);
        	p2.setAlignment(ParagraphAlignment.LEFT);
    
        	XWPFRun r1 = p1.createRun();

        	String t1 = "For Partnership Firm";
        	r1.setText(t1);
        	r1.setBold(true);
        	r1=p2.createRun();
        	r1.addCarriageReturn();
        	r1.addCarriageReturn();
        	r1.addCarriageReturn();
        	r1.addCarriageReturn();
        	r1.addCarriageReturn();
        	r1.addCarriageReturn();
        	r1.setFontFamily("Times New Roman");
        	
        
        	
             XWPFParagraph paragraph1 = document.createParagraph();
             r1.setText("Rs..............");
 
             r1.addTab();
             r1.addTab();
             r1.addTab();
             r1.addTab();
             r1.addTab();
             r1.addTab();
             r1.addTab();
             r1.addTab();
             r1.addTab();
             r1.addTab();
         
             String t2="\n\nPlace:........";
             r1.setText(t2);
             r1.addCarriageReturn();
             r1.addTab();
             r1.addTab();
             r1.addTab();
             r1.addTab();
             r1.addTab();
             r1.addTab();
             r1.addTab();
             r1.addTab();
             r1.addTab();
             r1.addTab();
             r1.addTab();
             
             String t3="\n\nDate:........";
             r1.setText(t3);
             r1.addCarriageReturn();

            BigInteger pos1 = BigInteger.valueOf(4500);
            setTabStop(paragraph1, STTabJc.Enum.forString("center"), pos1);
            BigInteger pos2 = BigInteger.valueOf(9000);
            setTabStop(paragraph1, STTabJc.Enum.forString("right"), pos2);
            
            BigInteger pos3 = BigInteger.valueOf(9000);
            setTabStop(paragraph1, STTabJc.Enum.forString("right"), pos3);
            
           
            r1.addCarriageReturn();
            XWPFParagraph paragraph = document.createParagraph();  
            
            paragraph = document.createParagraph(); 
            r1.setText("On Demand We, .................... ......................................................................."
            		+ ".........................................................................................................................."
            		+ "...............................................................................................................jointly and severely  promise to pay BANK OF BARODA or order at their office in.......................the sum of rupees..................... for value received, with interest thereon at the rate of ................% over Prime lending Rate of the Bank per annum with *monthly/Quarterly/half-yearly/yearly rests." );
            	
      
       
  XWPFParagraph p3 = document.createParagraph();  
            // 
 p3 = document.createParagraph(); 
	 r1 = p3.createRun();
 r1.setBold(true);
            String line1a="Personal Signatures of Partners";
            r1.setText(line1a);
     
            r1.addTab();
            r1.addTab();
            r1.addTab();
            r1.addTab();
            r1.addTab();
            r1.addTab();
            r1.addTab();
            String line1b="Firm Signature over the";
            r1.setText(line1b);
            r1.addCarriageReturn();
            r1.setText("In Full and without revenue stamp");
            r1.addTab();
            r1.addTab();
            r1.addTab();
            r1.addTab();
            r1.addTab();
            r1.addTab();
     
            String line2b="Stamp of appropriate";
            r1.setText(line2b);
            
            

           // File f = File.createTempFile("poi", ".docx");
            File f=new File("F:\\document\\partnership.docx");
            try (FileOutputStream fo = new FileOutputStream(f)) {
                document.write(fo);
            }
           // Desktop.getDesktop().open(f);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void setTabStop(XWPFParagraph oParagraph, STTabJc.Enum oSTTabJc, BigInteger oPos) {
        CTP oCTP = oParagraph.getCTP();
        CTPPr oPPr = oCTP.getPPr();
        if (oPPr == null) {
            oPPr = oCTP.addNewPPr();
        }

        CTTabs oTabs = oPPr.getTabs();
        if (oTabs == null) {
            oTabs = oPPr.addNewTabs();
        } 

        CTTabStop oTabStop = oTabs.addNewTab();
        oTabStop.setVal(oSTTabJc);
        oTabStop.setPos(oPos);
    }
}
