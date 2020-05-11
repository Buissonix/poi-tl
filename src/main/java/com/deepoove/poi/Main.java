package com.deepoove.poi;

import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.util.BytePictureUtils;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.util.HashMap;

public class Main {

	public void convertToPDF(String docPath, String pdfPath) {
		try {
			InputStream doc = new FileInputStream(new File(docPath));
			XWPFDocument document = new XWPFDocument(doc);
			PdfOptions options = PdfOptions.create();
			OutputStream out = new FileOutputStream(new File(pdfPath));
			PdfConverter.getInstance().convert(document, out, options);
			System.out.println("Done");
		} catch (FileNotFoundException ex) {
			System.out.println(ex.getMessage());
		} catch (IOException ex) {

			System.out.println(ex.getMessage());
		}
	}
	public static void main(String[] args) throws IOException {

		// LIEN DU TUTO EN ANGLAIS
		// https://github.com/Sayi/poi-tl/wiki/2.English-tutorial

		final File OUTPUT_DOCX = new File("out_docx.docx");
		final File OUTPUT_PDF = new File("out_pdf.pdf");

		//Supprimer l'output file
		//System.out.println("File exists : " + OUTPUT_DOCX.exists());
		if (OUTPUT_DOCX.exists()) {
			try{
				OUTPUT_DOCX.delete();}
			catch (Exception e){
				System.err.println("Le fichier est probablement ouvert dans un autre programme. Fermez-le et ressayez.");
				System.err.println(e.getMessage());
				e.printStackTrace();
			}
		}

		// Remplir un template déjà existant au format .docx
		XWPFTemplate template = XWPFTemplate.compile("template.docx").render(new HashMap<String, Object>(){{
			put("replaceMe", "Hello World! Check mon CV");
			put("pic", new PictureRenderData(100, 100, ".png", BytePictureUtils.getUrlByteArray("https://secure.gravatar.com/avatar/e04f14dbb6b83f8c9c36399f7136651b.png")));
			// remplacer le reste des valeurs {{valeur}} du template par ce qu'on veut ici

		}});
		FileOutputStream out = new FileOutputStream("out_docx.docx");
		template.write(out);
		out.flush();
		out.close();
		template.close();

		// Convertir un .docx en PDF
		// https://stackoverflow.com/questions/51440312/docx-to-pdf-converter-in-java

		InputStream in = new FileInputStream(OUTPUT_DOCX);
		XWPFDocument document = new XWPFDocument(in);
		PdfOptions options = PdfOptions.create();
		OutputStream pdf = new FileOutputStream(OUTPUT_PDF);
		PdfConverter.getInstance().convert(document, pdf, options);

		document.close();
		out.close();
	}


}

