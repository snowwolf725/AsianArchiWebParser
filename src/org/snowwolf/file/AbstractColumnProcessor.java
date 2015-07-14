package org.snowwolf.file;

import java.util.List;

import org.snowwolf.parser.PrintableDataItem;

public abstract class AbstractColumnProcessor {
	
	public abstract void process(List<PrintableDataItem> _items);

}
