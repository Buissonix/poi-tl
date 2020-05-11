package com.deepoove.poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

public class Main {
	public static void main(String[] args) throws IOException {

		// Remplir un template déjà existant au format .docx

		XWPFTemplate template = XWPFTemplate.compile("template.docx").render(new HashMap<String, Object>(){{
			put("replaceMe", "Hello World! Check mon CV");
			// remplacer le reste des valeurs {{valeur}} du template par ce qu'on veut ici

		}});
		FileOutputStream out = new FileOutputStream("out_template.docx");
		template.write(out);
		out.flush();
		out.close();
		template.close();

		// Convertir de .docx en PDF


	}
}
