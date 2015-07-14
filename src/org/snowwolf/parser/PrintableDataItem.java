package org.snowwolf.parser;


public class PrintableDataItem {
	
	private int m_row;
	
	private int m_col;
	
	private String m_value;
	
	public PrintableDataItem(int _row, int _col, String _value) {
		m_row = _row;
		m_col = _col;
		m_value = _value;
	}

	public int getRow() {
		return m_row;
	}

	public void setRow(int row) {
		this.m_row = row;
	}

	public int getCol() {
		return m_col;
	}

	public void setCol(int col) {
		this.m_col = col;
	}

	public String getValue() {
		return m_value;
	}

	public void setValue(String value) {
		this.m_value = value;
	}
}
