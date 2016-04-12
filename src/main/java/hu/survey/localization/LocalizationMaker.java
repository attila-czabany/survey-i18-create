package hu.survey.localization;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.commons.io.Charsets;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.WildcardFileFilter;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.json.JSONObject;

public class LocalizationMaker {
	private Collection<File> listFilesForFolder(final File folder) {
		return FileUtils.listFiles(folder, new WildcardFileFilter("*.xls"),
				null);
	}

	public static void main(String args[]) {
		LocalizationMaker lm = new LocalizationMaker();
		lm.makeLocalization();
	}

	private void makeLocalization() {
		final File folder = new File(System.getProperty("user.dir"));
		Collection<File> listFilesForFolder = listFilesForFolder(folder);
		Map<String, Map<String, String>> localization = new LinkedHashMap<String, Map<String, String>>();
		for (File file : listFilesForFolder) {
			System.out.println("Parsing file:  " + file.getAbsolutePath());
			parseFile(file, localization);
		}
		Set<Entry<String, Map<String, String>>> entrySet = localization
				.entrySet();
		for (Entry<String, Map<String, String>> entry : entrySet) {
			System.out.println("Creating file: " + entry.getKey());
			String content = new JSONObject(entry.getValue()).toString();
			try (Writer out = new BufferedWriter(new OutputStreamWriter(
					new FileOutputStream(entry.getKey()), Charsets.UTF_8));) {
				out.write(content);
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	private void parseFile(File file,
			Map<String, Map<String, String>> localization) {
		try(FileInputStream fis = new FileInputStream(file)) {
			POIFSFileSystem fs = new POIFSFileSystem(fis);
			HSSFWorkbook wb = new HSSFWorkbook(fs);
			HSSFSheet sheet = wb.getSheetAt(0);
			HSSFRow row = sheet.getRow(0);
			int i = 1;
			while (row.getCell(i) != null) {
				String language = row.getCell(i).getStringCellValue();
				Map<String, String> mapping = new LinkedHashMap<String, String>();
				int j = 1;
				while (sheet.getRow(j) != null) {
					extractTranslation(sheet, i, mapping, j);
					j++;
				}
				Map<String, String> map = localization.get(language);
				if (map == null) {
					localization.put(language, mapping);
				} else {
					map.putAll(map);
				}
				i++;
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void extractTranslation(HSSFSheet sheet, int cellNumber, Map<String, String> mapping,
			int rowNumber) {
		String key = sheet.getRow(rowNumber).getCell(0)
				.getStringCellValue();
		String value = sheet.getRow(rowNumber).getCell(cellNumber)
				.getStringCellValue();
		mapping.put(key, value);
	}

}
