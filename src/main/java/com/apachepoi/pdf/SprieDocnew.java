package com.apachepoi.pdf;

import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.TextSelection;
import com.spire.doc.fields.DocPicture;
import com.spire.doc.fields.TextRange;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

public class SprieDocnew {

	public static void main(String[] args) throws Exception {

		Document document = new Document("D:/EMPLOYEEDETAIL.docx");
		System.out.println(" Load the template document");

		Section section = document.getSections().get(0);
		System.out.println("Get the first section");

		Table table = section.getTables().get(0);
		System.out.println("Get the first table in the section");

		Map<String, Object> text = new HashMap<String, Object>();
		Map<String, String> newMap = new HashMap<String, String>();
		text.put("firstName", "Alex");
		text.put("lastName", "Anderson");
		List<String> st = new ArrayList<String>();
		st.add("+918009000909\n");
		st.add("+919878567670\n");
		st.add("+919887675644");
		text.put("mobilePhone", st);
		text.put("gender", "Male");
		text.put("email", "alexcurg@gmail.com");
		text.put("homeAddress", "4th cross San frnacico USA");
		text.put("dateOfBirth", "10-06-1995");

		for (Map.Entry<String, Object> entry : text.entrySet()) {
			try {
				Object tempValue = entry.getValue();
				String.valueOf(tempValue);
				if (tempValue instanceof List) {

					List<String> tempList = (List<String>) tempValue;
					String string = "";
					for (int i = 0; i < tempList.size(); i++) {
						if (i <= 0)
							string = tempList.get(i);
						else
							string = string + " " + tempList.get(i);
						

					}
					newMap.put(entry.getKey(), string);
				} else {
					newMap.put(entry.getKey(), String.valueOf(tempValue));

				}

			} catch (ClassCastException e) {
				System.out.println("ERROR: " + entry.getKey() + " -> " + entry.getValue() + " not added, as "
						+ entry.getValue() + " is not a String");

			}

			replaceTextinTable(newMap, table);
			System.out.println("Call the replaceTextinTable method to replace text in table");

			// replaceTextWithImage(document, "avatar", "D:/card.jpg");
			System.out.println("Call the replaceTextWithImage method to replace text with image");

			document.saveToFile("D:/MySpirDocx3.docx", FileFormat.Docx_2013);
			System.out.println("Save the result document");

			Document doc2 = new Document();
			System.out.println("Docoment instance creating");

			doc2.loadFromFile("D:/MySpirDocx3.docx");
			System.out.println("docoment is loaded");

			ToPdfParameterList ppl = new ToPdfParameterList();
			System.out.println("pdf instance ir creating");

			ppl.isEmbeddedAllFonts(true);
			System.out.println("embeded all fonts");

			ppl.setDisableLink(true);
			System.out.println("disabled hyperlink");

			doc2.setJPEGQuality(40);
			System.out.println("setting the quality");

			doc2.saveToFile("D:/MySpirDocTopdf.pdf", ppl);
			System.out.println("pdf is generated");

		}
	}

	// Replace text in table
	@SuppressWarnings("unchecked")
	static void replaceTextinTable(Map<String, String> newMap, Table table) {
		for (TableRow row : (Iterable<TableRow>) table.getRows()) {
			for (TableCell cell : (Iterable<TableCell>) row.getCells()) {
				for (Paragraph para : (Iterable<Paragraph>) cell.getParagraphs()) {
					for (Entry<String, String> entry : newMap.entrySet()) {

						para.replace("${" + entry.getKey() + "}", entry.getValue(), false, true);

					}
				}
			}
		}
	}

	// Replace text with image
	static void replaceTextWithImage(Document document, String stringToReplace, String imagePath)
			throws NullPointerException {
		TextSelection[] selections = document.findAllString("${" + stringToReplace + "}", false, true);
		int index = 0;
		TextRange range = null;
		for (Object obj1 : selections) {
			TextSelection textSelection = (TextSelection) obj1; // Creates a text selection for the given range.
			DocPicture pic = new DocPicture(document); // Initializes a new instance of the DocPicture class.
			pic.loadImage(imagePath); // image is loading
			pic.setWidth(160);
			pic.setHeight(120);
			range = textSelection.getAsOneRange();
			index = range.getOwnerParagraph().getChildObjects().indexOf(range);
			range.getOwnerParagraph().getChildObjects().insert(index, pic);
			range.getOwnerParagraph().getChildObjects().remove(range);
		}

	}

	@SuppressWarnings("unchecked")
	static void replaceTextinDocumentBody(Map<String, String> newMap, Document document) {
		for (Section section : (Iterable<Section>) document.getSections()) {
			for (Paragraph para : (Iterable<Paragraph>) section.getParagraphs()) {
				for (Map.Entry<String, String> entry : newMap.entrySet()) {
					para.replace("${" + entry.getKey() + "}", entry.getValue(), false, true);
				}
			}
		}
	}

	// Replace text in header or footer
	@SuppressWarnings("unchecked")
	static void replaceTextinHeaderorFooter(Map<String, String> data, HeaderFooter headerFooter) {
		for (Paragraph para : (Iterable<Paragraph>) headerFooter.getParagraphs()) {
			for (Map.Entry<String, String> entry : data.entrySet()) {
				para.replace("${" + entry.getKey() + "}", entry.getValue(), false, true);
			}
		}
	}

}
