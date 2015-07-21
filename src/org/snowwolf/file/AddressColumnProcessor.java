package org.snowwolf.file;

import java.io.FileNotFoundException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.snowwolf.parser.PrintableDataItem;

public class AddressColumnProcessor extends AbstractColumnProcessor {
	
	private static final int ADDRESS_COLUMN = 8;

	public AddressColumnProcessor() {
	}

	@Override
	public void process(List<PrintableDataItem> _items) {
		List<PrintableDataItem> zipCodes = new ArrayList<PrintableDataItem>();
		List<Integer> rows = new ArrayList<Integer>();
		List<String> nos = new ArrayList<String>();
		// span a column
		for(PrintableDataItem item : _items) {
			if(item.getCol() > ADDRESS_COLUMN) {
				item.setCol(item.getCol() + 1);
			} else if(item.getCol() == ADDRESS_COLUMN) {
				if(item.getValue().trim().equals("")) {
					rows.add(item.getRow());
				}
				PrintableDataItem zipCode = new PrintableDataItem(item.getRow(), item.getCol(), "");
				Pattern pattern = Pattern.compile("(\\d+)(.*)");
				Matcher matcher = pattern.matcher(item.getValue());
				if(matcher.find()) {
					zipCode.setValue(matcher.group(1));
					item.setValue(matcher.group(2));
				}
				item.setCol(ADDRESS_COLUMN + 1);
				zipCodes.add(zipCode);
			}
		}
		PrintWriter writer;
		try {
			writer = new PrintWriter("the-file-name.txt", "UTF-8");
			for(PrintableDataItem item : _items) {
				for(int row : rows) {
					if(item.getRow() == row && item.getCol() == 0) {
						nos.add(item.getValue());
						writer.println(item.getValue());
					}
				}
			}
			writer.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		
		System.out.println(nos.size());
		_items.addAll(zipCodes);
	}

}
