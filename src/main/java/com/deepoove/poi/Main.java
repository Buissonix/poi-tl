package com.deepoove.poi;

import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.util.BytePictureUtils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

public class Main {
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


	}
}
