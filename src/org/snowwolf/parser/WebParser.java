package org.snowwolf.parser;

import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class WebParser {
	
	private Map<String, Company> m_map_companies = new HashMap<String, Company>();
	
	private boolean m_isFail = false;

//	private int m_count;
	
	public static void main(String[] args) {
		WebParser parser = new WebParser();
		parser.downloadAll();
		
//		ExcelFileMerger merger = new ExcelFileMerger();
//		merger.addColumnProcesser(new AddressColumnProcessor());
//		merger.setPrintInfo("CompanyList_withZip", "CompanyList", 603);
//		merger.merge();
	}
	
//	private void test() {
//		Company cmp = new Company();
//		cmp.setNo("14877");
//		m_map_companies.put("14877", cmp);
//		getCompanyInfo();
//	}
//	
//	private void test2() {
//		Path path = Paths.get("nos.txt");
//		try {
//            Stream<String> lines = Files.lines(path);
//            m_count = 0;
//            lines.forEach(no -> {
//            	m_count++;
//        		Company cmp = new Company();
//        		cmp.setNo(no);
//        		m_map_companies.put(no, cmp);
//        		if(m_count % 200 == 0) {
//        			getCompanyInfo();
//        			saveExcelFile(m_count);
//        			m_map_companies.clear();
//        		}
//            });
//            getCompanyInfo();
//			saveExcelFile(m_count);
//        } catch (IOException ex) {
//
//        }
//	}
	
	private void downloadAll() {
		for(int i = 1;i <= 1000;i++) {
			try {
				m_isFail = false;
				m_map_companies.clear();
				getCompanyList(i);
				if(m_isFail == false) {
					getCompanyInfo();
				}
				if(m_isFail == false && i % 1 == 0) {
					saveExcelFile(i);
				}
			
				Thread.sleep(20000);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	private void getCompanyList(int _page) {
		try {
			String str_url = "http://www.asian-archi.com.tw/pages/yel_pag/search.asp?page=" + _page;
			URL url = new URL(str_url);
			Document xmlDoc = Jsoup.parse(url.openStream(), "MS950", url.toString());
			Elements ele_divTags = xmlDoc.getElementsByTag("div");
			for(Element ele_divTag : ele_divTags) {
				if(ele_divTag.attr("class").equals("custlft2") == false) {
					continue;
				}
				Elements ele_aTags = ele_divTag.getElementsByTag("a");
				for(Element ele_aTag : ele_aTags) {
					if(ele_aTag.attr("title").equals("") || ele_aTag.attr("href").contains("company") == false) {
						continue;
					}
					Company cmp = new Company();
					String name = ele_aTag.text().replaceAll(" \\- .*", "").trim();
					String no = ele_aTag.attr("href").replaceAll("/company/", "").replaceAll("/index.html", "").trim();
					cmp.setNo(no);
					cmp.setName(name);
					m_map_companies.put(no, cmp);
					System.out.println("No:" + no + "\tName:" + name);
				}
			}
		} catch (MalformedURLException e) {
			e.printStackTrace();
			m_isFail = true;
		} catch (IOException e) {
			e.printStackTrace();
			m_isFail = true;
		}
	}
	
//	private void getCompanyList2(int _page) {
//		try {
//			String str_url = "http://www.archi.net.tw/html/webcaseadvclsd";
//			if(_page != 1) {
//				str_url += _page;
//			}
//			str_url += ".asp";
//			URL url = new URL(str_url);
//			Document xmlDoc = Jsoup.parse(url.openStream(), "MS950", url.toString());
//			Elements liTags = xmlDoc.getElementsByTag("li");
//			for(Element li : liTags) {
//				Elements ele_aTags = li.getElementsByTag("a");
//				for(Element ele_aTag : ele_aTags) {
//					Company cmp = new Company();
//					String name = ele_aTag.text().replaceAll(" \\- .*", "").trim();
//					String no = ele_aTag.attr("href").replaceAll("/company/", "").replaceAll("/index.html", "").trim();
//					cmp.setNo(no);
//					cmp.setName(name);
//					m_map_companies.put(no, cmp);
//					System.out.println("No:" + no + "\tName:" + name);
//				}
//			}
//		} catch (MalformedURLException e) {
//			e.printStackTrace();
//			m_isFail = true;
//		} catch (IOException e) {
//			e.printStackTrace();
//			m_isFail = true;
//		}
//	}

	private void getCompanyInfo() {
		try {
			for(String key : m_map_companies.keySet()) {
				Thread.sleep(4000);
				Company cmp = m_map_companies.get(key);
				URL url = new URL("http://www.archi.net.tw/company/" + key + "/index.html");
				Document xmlDoc =  Jsoup.parse(url.openStream(), "MS950", url.toString());
				Elements ele_h4s = xmlDoc.getElementsByTag("h4");
				for(Element ele_h4 : ele_h4s) {
					if(ele_h4.text().equals("相關資料")) {
						cmp.setRelateData(ele_h4.nextElementSibling().text());
					}
				}
				Elements divs = xmlDoc.getElementsByTag("div");
				for(Element div : divs) {
					if(div.attr("id").equals("divcompname")) {
						Elements ele_h1s = div.getElementsByTag("h1");
						for(Element ele_h1 : ele_h1s) {
							cmp.setName(ele_h1.text());
						}
						continue;
					} else if(div.attr("id").equals("custsata01") == false) {
						continue;
					}
					
					Elements ele_lis = div.getElementsByTag("li");
					if(div.getElementsByTag("ul").size() != 0) {
						ele_lis = div.getElementsByTag("ul").get(0).getElementsByTag("li");
					}
					for(Element ele_li : ele_lis) {
						String tag = ele_li.child(0).text().trim();
						String value = ele_li.textNodes().get(0).text().replaceAll("\\&nbsp;", "").trim();
						if(ele_li.childNode(1) instanceof Element) {
							value = ele_li.child(1).text().replaceAll("\\&nbsp;", "").trim();
						}
						System.out.println(tag + "\t" + value);
						switch(tag) {
						case "會員編號":
//							cmp.setNo(value);
							break;
						case "公司英文":
							cmp.setEngName(value);
							break;
						case "統一編號":
							cmp.setUniNo(value);
							break;
						case "公會社團":
							cmp.setClub(value);
							break;
						case "聯絡電話":
							cmp.setTel(value);
							break;
						case "聯絡傳真":
							cmp.setFax(value);
							break;
						case "Skype":
							cmp.setSkype(value);
							break;
						case "通訊地址":
							Pattern pattern = Pattern.compile("(\\d+)(.*)");
							Matcher matcher = pattern.matcher(value);
							String str_zip = "";
							if(matcher.find()) {
								str_zip = matcher.group(1);
								cmp.setAddress(matcher.group(2));
								cmp.setZip(Integer.parseInt(str_zip));
							} else {
								cmp.setAddress(value);
							}
							break;
						case "E-mail":
							value = value.replaceAll("不公佈或未填寫", "");
							cmp.setEmail(value);
							break;
						case "所屬分類":
							cmp.setCategory(value);
							break;
						case "經營型態":
							cmp.setType(value);
							break;
						case "服務地區":
							cmp.setServiceAddress(value);
							break;
						case "公司網址":
							cmp.setWeb(value);
							break;
						case "聯絡手機":
							cmp.setPhone(value);
							break;
						default:
							break;
						}
					}
				}
			}
		} catch (MalformedURLException e) {
			e.printStackTrace();
			m_isFail = true;
		} catch (IOException e) {
			e.printStackTrace();
			m_isFail = true;
		} catch (InterruptedException e) {
			e.printStackTrace();
			m_isFail = true;
		}
	}

	public void saveExcelFile(int _page) {
		List<PrintableDataItem> items = new ArrayList<PrintableDataItem>();
		PrintableDataItem item = null;
		int rowIndex = 1;
		for(String key : m_map_companies.keySet()) {
			Company company = m_map_companies.get(key);
			item = new PrintableDataItem(rowIndex, 0, company.getNo());
			items.add(item);
			item = new PrintableDataItem(rowIndex, 1, company.getName());
			items.add(item);
			item = new PrintableDataItem(rowIndex, 2, company.getEngName());
			items.add(item);
			item = new PrintableDataItem(rowIndex, 3, company.getUniNo());
			items.add(item);
			item = new PrintableDataItem(rowIndex, 4, company.getClub());
			items.add(item);
			item = new PrintableDataItem(rowIndex, 5, company.getPhone());
			items.add(item);
			item = new PrintableDataItem(rowIndex, 6, company.getTel());
			items.add(item);
			item = new PrintableDataItem(rowIndex, 7, company.getFax());
			items.add(item);
			item = new PrintableDataItem(rowIndex, 8, company.getSkype());
			items.add(item);
			item = new PrintableDataItem(rowIndex, 9, company.getZip() + "");
			items.add(item);
			item = new PrintableDataItem(rowIndex, 10, company.getAddress());
			items.add(item);
			item = new PrintableDataItem(rowIndex, 11, company.getEmail());
			items.add(item);
			item = new PrintableDataItem(rowIndex, 12, company.getCategory());
			items.add(item);
			item = new PrintableDataItem(rowIndex, 13, company.getType());
			items.add(item);
			item = new PrintableDataItem(rowIndex, 14, company.getServiceAddress());
			items.add(item);
			item = new PrintableDataItem(rowIndex, 15, company.getWeb());
			items.add(item);
			item = new PrintableDataItem(rowIndex, 16, company.getRelateData());
			items.add(item);
			rowIndex++;
		}
		ReportGenerator generator = new ReportGenerator();
		generator.setPrintInfo("CompanyList", "CompanyList_" + _page, items);
		generator.genReport();
	}
}
