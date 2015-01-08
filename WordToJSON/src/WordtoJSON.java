import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

public class WordtoJSON {
	public static void main(String[] args) {
		try {
			new WordtoJSON().process("C:/001.doc");
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	@SuppressWarnings("unchecked")
	private void process(String path) throws IOException {
		FileInputStream fileInputStream = new FileInputStream(path);
		HWPFDocument document = new HWPFDocument(fileInputStream);
		int text = document.getRange().numCharacterRuns();

		JSONArray mainArray = new JSONArray();
		JSONObject mainJsonObject = new JSONObject();
		mainArray.add(mainJsonObject);

		JSONArray headerArr = new JSONArray();
		mainJsonObject.put("header", headerArr);
		JSONArray bodyArr = new JSONArray();
		mainJsonObject.put("body", bodyArr);
		for (int i = 0; i < text; i++) {
			CharacterRun characterRun = document.getRange().getCharacterRun(i);
			System.out.println(characterRun.text());
			System.out.println("color " + characterRun.getColor());
			System.out.println("bold " + characterRun.isBold());
			System.out.println("font size " + (characterRun.getFontSize() / 2));
			System.out.println();

			if (i <= 1) {
				JSONObject header = new JSONObject();
				header.put("text", characterRun.text());
				header.put("fontSize", characterRun.getFontSize() / 2);
				header.put("fontName", characterRun.getFontName());
				header.put("textColor", characterRun.getColor() == 6 ? "#ff0000" : "#000000");
				header.put("textAlign", "center");
				headerArr.add(header);
				continue;
			}

			JSONObject obj = new JSONObject();
			obj.put("text", characterRun.text());
			obj.put("fontSize", characterRun.getFontSize() / 2);
			obj.put("fontName", "GE_SS_Two");
			obj.put("textColor", characterRun.getColor() == 6 ? "#ff0000" : "#000000");
			obj.put("textAlign", "right");
			bodyArr.add(obj);
		}
		FileWriter file = new FileWriter("c:\\001.json");
		file.write(mainArray.toJSONString());
		file.flush();
		file.close();
	}
}
