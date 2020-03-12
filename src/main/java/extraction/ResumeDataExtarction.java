package extraction;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
/*
 * anil Jm @8897444649
 */
public class ResumeDataExtarction {
	static JSONArray list_array = new JSONArray();
	static JSONObject keyvalue_data = new JSONObject();
	static JSONObject main_extarted_data = new JSONObject();

	static /*
			 * tables
			 */
	JSONArray Table_number = new JSONArray();
	static JSONArray Table_rows_data;
	static JSONObject table_main = new JSONObject();
	static JSONObject rows;

	@SuppressWarnings("unchecked")
	public static void main(String[] args) {

		try {
			/* change folder path here */

			File folder = new File(
					"C:\\Users\\anilj\\Desktop\\Machine learning_projects\\ResumeExtraction\\word-doc-parser\\input");
			File[] listOfFiles = folder.listFiles();

			for (File file : listOfFiles) {
				list_array = new JSONArray();
				keyvalue_data = new JSONObject();
				main_extarted_data = new JSONObject();

				/*
				 * tables
				 */
				Table_number = new JSONArray();
				Table_rows_data = new JSONArray();
				table_main = new JSONObject();
				rows = new JSONObject();

				if (file.isFile()) {
					System.out.println(file.getName());
					System.out.println(file.getAbsolutePath());

					FileInputStream fis = new FileInputStream(file.getAbsolutePath());
					XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(fis));

					getParagaraphData(xdoc);
					gettablesData(xdoc);

					main_extarted_data.put("tables", table_main);
					createJsonfile(main_extarted_data, file.getName());
				}

			}

		} catch (Exception ex) {
			ex.printStackTrace();
		}

	}

	@SuppressWarnings("unchecked")
	private static void gettablesData(XWPFDocument xdoc) {

		try {

			Iterator<IBodyElement> bodyElementIterator = xdoc.getBodyElementsIterator();
			while (bodyElementIterator.hasNext()) {
				IBodyElement element = bodyElementIterator.next();

				if ("TABLE".equalsIgnoreCase(element.getElementType().name())) {
					List<XWPFTable> tableList = element.getBody().getTables();
					int k = 0;
					for (XWPFTable table : tableList) {
						System.out.println("Total Number of Rows of Table:" + table.getNumberOfRows());
						rows = new JSONObject();
						for (int i = 0; i < table.getRows().size(); i++) {
							Table_rows_data = new JSONArray();
							for (int j = 0; j < table.getRow(i).getTableCells().size(); j++) {
								System.out.println(table.getRow(i).getCell(j).getText());
								Table_rows_data.add(table.getRow(i).getCell(j).getText());
							}
							rows.put(i, Table_rows_data);
						}

						table_main.put(k, rows);
						k++;
					}
				}
			}
			System.out.println(table_main.toString());

		} catch (Exception ex) {
			ex.printStackTrace();
		}

	}

	@SuppressWarnings("unchecked")
	private static void getParagaraphData(XWPFDocument xdoc) {
		// TODO Auto-generated method stub
		List<XWPFParagraph> paragraphList = xdoc.getParagraphs();

		for (XWPFParagraph paragraph : paragraphList) {
			if (!paragraph.getText().isEmpty() && paragraph.getText().length() != 0) {
				if (paragraph.getText().contains(":")) {
					String keypairs[] = paragraph.getText().split(":");
					if (keypairs.length != 0 && keypairs.length == 2) {
						keyvalue_data.put(keypairs[0], keypairs[1]);

					} else {
						String a = paragraph.getText().trim();
						list_array.add(a);
					}
				} else {
					String a = paragraph.getText().trim();
					list_array.add(a);
				}

			}
			System.out.println(paragraph.getText());
			System.out.println(paragraph.getAlignment());// LEFT
			System.out.print(paragraph.getRuns().size());
			System.out.println(paragraph.getStyle());

			// Returns numbering format for this paragraph, eg bullet or
			// lowerLetter.
			System.out.println(paragraph.getNumFmt());
			System.out.println(paragraph.getAlignment());// LEFT

			System.out.println(paragraph.isWordWrapped());

			System.out.println("********************************************************************");

			for (XWPFRun rn : paragraph.getRuns()) {

				System.out.println(rn.isBold());
				System.out.println(rn.isHighlighted());
				System.out.println(rn.isCapitalized());
				System.out.println(rn.getFontSize());
			}

		}

		main_extarted_data.put("key_value_data", keyvalue_data);
		main_extarted_data.put("paragarphs_data", list_array);

	}

	private static void createJsonfile(JSONObject h, String filename) {

		// Write JSON file
		try {
			System.out.println("filename" + filename + ".json");

			/*
			 * change file path here
			 */
			FileWriter file = new FileWriter(
					"C:\\Users\\anilj\\Desktop\\Machine learning_projects\\ResumeExtraction\\word-doc-parser\\output\\"
							+ filename + ".json");
			file.write(main_extarted_data.toJSONString());
			file.close();

		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
