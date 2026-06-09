package Debtors;

// last updated 10.05.26 @ 20.30 hrs
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;
import java.util.concurrent.TimeUnit;
import java.util.stream.Collectors;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderExtent;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class ledger {
	public String co;
	public int code;
	public int bucket;
	public double amount;

	public int getBucket() {
		return bucket;
	}

	public void setBucket(int bucket) {
		this.bucket = bucket;
	}

	public String getCo() {
		return co;
	}

	public void setCo(String co) {
		this.co = co;
	}

	public int getCode() {
		return code;
	}

	public void setCode(int code) {
		this.code = code;
	}

	public double getAmount() {
		return amount;
	}

	public void setAmount(double amount) {
		this.amount = amount;
	}

	public ledger(String co, int code, int bucket, double amount) {
		this.co = co;
		this.code = code;
		this.bucket = bucket;
		this.amount = amount;
	}
}

class master {
	public int code;
	public String name;
	public String status;
	public String group;
	public String panno;
	public double deposit;
	public double ecl;
	public int days;

	public master(int code, String status, String name, String group, String panno, double deposit, double ecl,
			int days) {
		super();
		this.code = code;
		this.status = status;
		this.name = name;
		this.group = group;
		this.panno = panno;
		this.deposit = deposit;
		this.ecl = ecl;
		this.days = days;
	}

	public int getCode() {
		return code;
	}

	public void setCode(int code) {
		this.code = code;
	}

	public String getStatus() {
		return status;
	}

	public void setStatus(String status) {
		this.status = status;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getGroup() {
		return group;
	}

	public void setGroup(String group) {
		this.group = group;
	}

	public String getPanno() {
		return panno;
	}

	public void setPanno(String panno) {
		this.panno = panno;
	}

	public double getDeposit() {
		return deposit;
	}

	public void setDeposit(double deposit) {
		this.deposit = deposit;
	}

	public double getEcl() {
		return ecl;
	}

	public void setEcl(double ecl) {
		this.ecl = ecl;
	}

	public int getDays() {
		return days;
	}

	public void setDays(int days) {
		this.days = days;
	}

	@Override
	public String toString() {
		return name + " ~" + status + "~" + group + "~ " + panno;
	}
}

public class ageing {
	public static void main(String[] args) throws IOException, Exception, InvalidFormatException {
		System.out
				.println("updated on " + LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd-MM-yy hh:mm:ss  "))
						+ System.getProperty("user.dir"));
		@SuppressWarnings("resource")
		XSSFWorkbook wb = new XSSFWorkbook();
		CellStyle Title = null;
		Font font = wb.createFont();
		font.setFontHeightInPoints((short) 12);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(true);
		font.setItalic(false);
		Title = wb.createCellStyle();
		Title.setFont(font);
		CellStyle header = null;
		font = wb.createFont();
		font.setFontHeightInPoints((short) 11);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.RED.getIndex());
		font.setBold(true);
		font.setItalic(false);
		header = wb.createCellStyle();
		header.setFont(font);
		header.setAlignment(HorizontalAlignment.LEFT);
		CellStyle header2 = null;
		font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(true);
		font.setItalic(false);
		header2 = wb.createCellStyle();
		header2.setFont(font);
		header2.setAlignment(HorizontalAlignment.LEFT);
		CellStyle body = null;
		font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(false);
		font.setItalic(false);
		body = wb.createCellStyle();
		body.setFont(font);
		CellStyle amount0 = wb.createCellStyle();
		DataFormat format = wb.createDataFormat();
		font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(false);
		font.setItalic(false);
		amount0 = wb.createCellStyle();
		amount0.setDataFormat(format.getFormat("#,##,##0"));
		amount0.setFont(font);
		CellStyle amount1 = wb.createCellStyle();
		format = wb.createDataFormat();
		font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(false);
		font.setItalic(false);
		amount1 = wb.createCellStyle();
		amount1.setDataFormat(format.getFormat("#,##,##0.00"));
		amount1.setFont(font);
		CellStyle amount2 = wb.createCellStyle();
		format = wb.createDataFormat();
		font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(true);
		font.setItalic(false);
		amount2 = wb.createCellStyle();
		amount2.setDataFormat(format.getFormat("#,##,##0.00"));
		amount2.setFont(font);
		CellStyle total = wb.createCellStyle();
		format = wb.createDataFormat();
		font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.RED.getIndex());
		font.setBold(true);
		font.setItalic(false);
		total = wb.createCellStyle();
		total.setDataFormat(format.getFormat("#,##,##0"));
		total.setFont(font);
		CellStyle rightAligned = wb.createCellStyle();
		font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.RED.getIndex());
		font.setBold(true);
		font.setItalic(false);
		rightAligned.setAlignment(HorizontalAlignment.RIGHT);
		rightAligned.setFont(font);
		CellStyle rightAligned2 = wb.createCellStyle();
		font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(true);
		font.setItalic(false);
		rightAligned2.setAlignment(HorizontalAlignment.RIGHT);
		rightAligned2.setFont(font);
		CellStyle rightAlignedbody = wb.createCellStyle();
		font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Calibri");
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(true);
		font.setItalic(false);
		rightAlignedbody.setAlignment(HorizontalAlignment.RIGHT);
		rightAlignedbody.setFont(font);
		Date start = new Date();
		String ddmmyy = null;
		String filepath = "D:/debtors/drs 2026-27/may 2026/";
		int days = 0;
		int days1 = 0;
		double uacbal = 0.00;
		double amount = 0;
		String tap[] = null;
		int row = 5;
		int col = 1;
		DateTimeFormatter ft = DateTimeFormatter.ofPattern("dd.MM.yyyy");
		String reportdt = "31.05.2026";
		// Aging bucket conditions
		int[][] bucket1 = { { 0, 30 }, { 31, 60 }, { 61, 90 }, { 91, 120 }, { 121, 150 }, { 151, 180 }, { 181, 9999 } };
		int[][] bucket2 = { { -99, 0 }, { 1, 180 }, { 181, 360 }, { 361, 720 }, { 721, 1080 }, { 1081, 2600 },
				{ 2601, 9999 } };
		String[][] buckethead = { { "<30 days", "Not due" }, { "31-60 days", "<Six Months" },
				{ "61-90 days", "6m-12 month" }, { "91-120 days", "13-24 monrh" }, { "121-150 days", "25-36 monrh" },
				{ "151-180 days", ">36 month" }, { ">181 days", ">61 months" } };
		TreeMap<Integer, Integer> agegroup = new TreeMap<Integer, Integer>();
		for (int a = 0; a < bucket1.length; a++)
			for (int sl = bucket1[a][0]; sl <= bucket1[a][1]; sl++)
				agegroup.put(sl, a);
		TreeMap<Integer, Integer> schedule3 = new TreeMap<Integer, Integer>();
		for (int a = 0; a < bucket2.length; a++)
			for (int sl = bucket2[a][0]; sl <= bucket2[a][1]; sl++)
				schedule3.put(sl, a);
		// uac clearing header
		int sl = 0;
		TreeMap<String, Integer> alpha = new TreeMap<String, Integer>();
		TreeMap<Integer, String> alphareverse = new TreeMap<Integer, String>();
		FileInputStream xlsfile = new FileInputStream(
				filepath.substring(0, filepath.lastIndexOf("/") - 8) + "master.xlsx");
		Workbook xls = WorkbookFactory.create(xlsfile);
		System.out.println("updating customer list  ...");
		// updating alpha keys
		sl = 600001;
		Sheet sht = xls.getSheetAt(1);
		for (Row r : sht) {
			if (r.getRowNum() > 0) {
				alpha.put(r.getCell(0).getStringCellValue(), sl);
				alphareverse.put(sl++, r.getCell(0).getStringCellValue());
			}
		}
		List<master> code = new ArrayList<master>();
		int c = 0;
		sl = 0;
		sht = xls.getSheetAt(0);
		for (Row r : sht) {
			if (r.getRowNum() > 0) {
				{
					switch (r.getCell(0).getCellTypeEnum()) {
					case STRING:
						sl = alpha.get(r.getCell(0).getStringCellValue());
						break;
					default:
						sl = (int) r.getCell(0).getNumericCellValue();
					}
					master a1 = new master(sl, r.getCell(1).getStringCellValue(), r.getCell(2).getStringCellValue(),
							r.getCell(3).getStringCellValue(), r.getCell(4).getStringCellValue(),
							r.getCell(5).getNumericCellValue(), r.getCell(6).getNumericCellValue(),
							(int) r.getCell(7).getNumericCellValue());
					code.add(a1);
					 System.out.println(c++);
				}
			}
		}
		Map<Integer, Integer> duedays = code.stream().collect(Collectors.toMap(master::getCode, master::getDays));
		// updating ledger
		sl = 0;
		c = 0;
		List<ledger> trans = new ArrayList<ledger>();
		List<ledger> sch3 = new ArrayList<ledger>();
		System.out.println("updating ledger");
		xlsfile = new FileInputStream(filepath + "ledger.xlsx");
		xls = WorkbookFactory.create(xlsfile);
		sht = xls.getSheetAt(0);
		for (Row r : sht) {
			if (r.getRowNum() > 0) {
				{
					switch (r.getCell(1).getCellTypeEnum()) {
					case STRING:
						sl = alpha.get(r.getCell(1).getStringCellValue().trim());
						break;
					default:
						sl = (int) r.getCell(1).getNumericCellValue();
					}
					ddmmyy = r.getCell(5).getStringCellValue();
					days = (int) ChronoUnit.DAYS.between(LocalDate.parse(ddmmyy, ft), LocalDate.parse(reportdt, ft));
					System.out.println(days);					
					days1 = (int) ChronoUnit.DAYS.between(LocalDate.parse(ddmmyy, ft).plusDays(duedays.get(sl)),
							LocalDate.parse(reportdt, ft));
					ledger a1 = new ledger(r.getCell(0).getStringCellValue().trim(), sl, agegroup.get(days),
							r.getCell(7).getNumericCellValue());
					ledger a2 = new ledger(r.getCell(0).getStringCellValue().trim(), sl, schedule3.get(days1),
							r.getCell(7).getNumericCellValue());
					System.out.println(c++ + "~" + sl );
					trans.add(a1);
					sch3.add(a2);
				}
			}
		}
		// calculation
		Set<Integer> activecodes = trans.stream().map(ledger::getCode).collect(Collectors.toCollection(TreeSet::new));
		Map<Integer, String> custcode = code.stream().collect(Collectors.toMap(master::getCode, master::toString));
		TreeMap<Integer, Double> ageing_credit = trans.stream().filter(a -> a.getAmount() < 0).collect(
				Collectors.groupingBy(ledger::getCode, TreeMap::new, Collectors.summingDouble(ledger::getAmount)));
		TreeMap<Integer, Double> sch3_credit = trans.stream().filter(a -> a.getAmount() < 0).collect(
				Collectors.groupingBy(ledger::getCode, TreeMap::new, Collectors.summingDouble(ledger::getAmount)));
		TreeMap<Integer, Map<Integer, Double>> ageing_debit = trans.stream().filter(a -> a.getAmount() > 0)
				.collect(Collectors.groupingBy(ledger::getCode, TreeMap::new,
						Collectors.groupingBy(ledger::getBucket, Collectors.summingDouble(ledger::getAmount))));
		TreeMap<Integer, Map<Integer, Double>> schedule3_debit = sch3.stream().filter(a -> a.getAmount() > 0)
				.collect(Collectors.groupingBy(ledger::getCode, TreeMap::new,
						Collectors.groupingBy(ledger::getBucket, Collectors.summingDouble(ledger::getAmount))));
		Map<Integer, Double> ecl = code.stream().collect(Collectors.toMap(master::getCode, master::getEcl));
		Map<Integer, Double> deposit = code.stream().collect(Collectors.toMap(master::getCode, master::getDeposit));
		// uac clearing for regular ageing
		for (Entry<Integer, Double> uac : ageing_credit.entrySet()) {
			uacbal = uac.getValue();
			{
				if (ageing_debit.get(uac.getKey()) != null) {
					for (int a = 6; a >= 0; a--) {
						if (ageing_debit.get(uac.getKey()).get(a) != null) {
							amount = ageing_debit.get(uac.getKey()).get(a);
							{
								if (amount + uacbal > 0) {
									amount += uacbal;
									uacbal = 0;
								} else {
									uacbal += amount;
									amount = 0;
								}
								ageing_debit.get(uac.getKey()).replace(a, amount);
							}
						}
					}
					ageing_credit.replace(uac.getKey(), uacbal);
				}
			}
		}
		// uac clearing for schedule 3 ageing
		for (Entry<Integer, Double> uac : sch3_credit.entrySet()) {
			uacbal = uac.getValue();
			{
				if (schedule3_debit.get(uac.getKey()) != null) {
					for (int a = 6; a >= 0; a--) {
						if (schedule3_debit.get(uac.getKey()).get(a) != null) {
							amount = schedule3_debit.get(uac.getKey()).get(a);
							{
								if (amount + uacbal > 0) {
									amount += uacbal;
									uacbal = 0;
								} else {
									uacbal += amount;
									amount = 0;
								}
								schedule3_debit.get(uac.getKey()).replace(a, amount);
							}
						}
					}
					sch3_credit.replace(uac.getKey(), uacbal);
				}
			}
		}
		// printing
		row = 5;
		col = 1;
		int temp = 0;
		XSSFSheet sht1 = wb.createSheet("Ageing");
		for (Integer s : activecodes) {
			sht1.createRow(row).createCell(col)
					.setCellValue((alphareverse.get(s) != null) ? (alphareverse.get(s)) : s.toString());
			sht1.getRow(row).getCell(col++).setCellStyle(body);
			tap = custcode.get(s).split("~");
			for (int a = 0; a < tap.length; a++) {
				sht1.getRow(row).createCell(col).setCellValue(tap[a]);
				sht1.getRow(row).getCell(col++).setCellStyle(body);
			}
			if (ageing_debit.get(s) != null) {
				for (int a = 6; a < 13; a++) {
					amount = ageing_debit.get(s).get(temp) == null ? 0 : ageing_debit.get(s).get(temp);
					sht1.getRow(row).createCell(a).setCellValue(amount);
					sht1.getRow(row).getCell(a).setCellStyle(amount0);
					temp++;
				}
				temp = 0;
			}
			if (ageing_credit.get(s) != null) {
				sht1.getRow(row).createCell(13).setCellValue(ageing_credit.get(s));
				sht1.getRow(row).getCell(13).setCellStyle(amount0);
			}
			sht1.getRow(row).createCell(14).setCellFormula("sum(g" + (row + 1) + ":n" + (row + 1) + ")");
			sht1.getRow(row).getCell(14).setCellStyle(amount0);
			sht1.getRow(row).createCell(15).setCellValue(deposit.get(s));
			sht1.getRow(row).getCell(15).setCellStyle(amount0);
			sht1.getRow(row).createCell(16).setCellValue(ecl.get(s));
			sht1.getRow(row).getCell(16).setCellStyle(amount0);
			col = 1;
			row++;
		}
		row = 1;
		col = 1;
		String[] co = trans.stream().map(ledger::getCo).collect(Collectors.toCollection(TreeSet::new))
				.toArray(String[]::new);	
		Map<String, String> cocode = new TreeMap<String, String>();
		cocode.put("ND00", "The Nilgiri Dairy Farm P Ltd");
		cocode.put("AN00", "Appu Nutritions Pvt Ltd");
		cocode.put("NF00", "Nilgiris Franchisee Ltd");
		cocode.put("NM00", "Nilgiris Machanised Bakery P Ltd");
		sht1.createRow(row).createCell(col).setCellValue(cocode.get(co[0]));
		sht1.getRow(row).getCell(col).setCellStyle(Title);
		row = 2;
		sht1.createRow(row).createCell(col).setCellValue("Debtors as on " + reportdt);
		sht1.getRow(row).getCell(col).setCellStyle(Title);
		row = 4;
		col = 6;
		sht1.createRow(row);
		for (int a = 6; a < buckethead.length + 6; a++) {
			sht1.getRow(row).createCell(col).setCellValue(buckethead[a - 6][0]);
			sht1.getRow(row).getCell(col).setCellStyle(rightAligned);
			sht1.setColumnWidth(col++, 2900);
		}
		sht1.getRow(row).createCell(col).setCellValue("Credit");
		sht1.getRow(row).getCell(col).setCellStyle(rightAligned);
		sht1.setColumnWidth(col++, 2900);
		sht1.getRow(row).createCell(col).setCellValue("Total");
		sht1.getRow(row).getCell(col).setCellStyle(rightAligned);
		sht1.setColumnWidth(col++, 2900);
		sht1.getRow(row).createCell(col).setCellValue("Deposit");
		sht1.getRow(row).getCell(col).setCellStyle(rightAligned);
		sht1.setColumnWidth(col++, 2900);
		sht1.getRow(row).createCell(col).setCellValue("ECL Provn");
		sht1.getRow(row).getCell(col).setCellStyle(rightAligned);
		sht1.setColumnWidth(col++, 2900);
		col = 1;
		sht1.getRow(row).createCell(col).setCellValue("Code");
		sht1.getRow(row).getCell(col).setCellStyle(header);
		sht1.setColumnWidth(col++, 1800);
		sht1.getRow(row).createCell(col).setCellValue("Customer");
		sht1.getRow(row).getCell(col).setCellStyle(header);
		sht1.setColumnWidth(col++, 9000);
		sht1.getRow(row).createCell(col).setCellValue("Status");
		sht1.getRow(row).getCell(col).setCellStyle(header);
		sht1.setColumnWidth(col++, 2600);
		sht1.getRow(row).createCell(col).setCellValue("CustGroup");
		sht1.getRow(row).getCell(col).setCellStyle(header);
		sht1.setColumnWidth(col++, 3200);
		sht1.getRow(row).createCell(col).setCellValue("Pan No.");
		sht1.getRow(row).getCell(col).setCellStyle(header);
		sht1.setColumnWidth(col++, 3100);
		sht1.setColumnWidth(0, 400);
		sht1.setDisplayGridlines(false);
		// vertical total
		sl = sht1.getLastRowNum();
		sht1.createRow(sl + 2);
		for (col = 6; col < 17; col++) {
			// System.out.println(col);
			sht1.getRow(sl + 2).createCell(col)
					.setCellFormula("sum(" + (char) (97 + col) + "6:" + (char) (97 + col) + (sl + 1) + ")");
			sht1.getRow(sl + 2).getCell(col).setCellStyle(total);
		}
		sl = sht1.getLastRowNum();
		sht1.getRow(sl).createCell(5).setCellValue("Total");
		sht1.getRow(sl).getCell(5).setCellStyle(header2);
		sl = sht1.getLastRowNum();
		PropertyTemplate pt = new PropertyTemplate();
		pt.drawBorders(new CellRangeAddress(4, sl, 1, sht1.getRow(4).getLastCellNum() - 1), BorderStyle.THIN,
				IndexedColors.GREY_25_PERCENT.getIndex(), BorderExtent.ALL);
		pt.applyBorders(sht1);
		// wb.cloneSheet(0); wb.setSheetName(1, "CreditLoss");
		// ********** schedule 3 format ************
		row = 5;
		col = 1;
		temp = 0;
		XSSFSheet sht2 = wb.createSheet("schedule_3");
		for (Integer s : activecodes) {
			sht2.createRow(row).createCell(col)
					.setCellValue((alphareverse.get(s) != null) ? (alphareverse.get(s)) : s.toString());
			sht2.getRow(row).getCell(col++).setCellStyle(body);
			tap = custcode.get(s).split("~");
			for (int a = 0; a < tap.length; a++) {
				sht2.getRow(row).createCell(col).setCellValue(tap[a]);
				sht2.getRow(row).getCell(col++).setCellStyle(body);
			}
			if (schedule3_debit.get(s) != null) {
				for (int a = 6; a < 13; a++) {
					amount = schedule3_debit.get(s).get(temp) == null ? 0 : schedule3_debit.get(s).get(temp);
					sht2.getRow(row).createCell(a).setCellValue(amount);
					sht2.getRow(row).getCell(a).setCellStyle(amount0);
					temp++;
				}
				temp = 0;
			}
			if (ageing_credit.get(s) != null) {
				sht2.getRow(row).createCell(13).setCellValue(sch3_credit.get(s));
				sht2.getRow(row).getCell(13).setCellStyle(amount0);
			}
			sht2.getRow(row).createCell(14).setCellFormula("sum(g" + (row + 1) + ":n" + (row + 1) + ")");
			sht2.getRow(row).getCell(14).setCellStyle(amount0);
			sht2.getRow(row).createCell(15).setCellValue(deposit.get(s));
			sht2.getRow(row).getCell(15).setCellStyle(amount0);
			sht2.getRow(row).createCell(16).setCellValue(ecl.get(s));
			sht2.getRow(row).getCell(16).setCellStyle(amount0);
			col = 1;
			row++;
		}
		row = 1;
		col = 1;
		sht2.createRow(row).createCell(col).setCellValue(cocode.get(co[0]));
		sht2.getRow(row).getCell(col).setCellStyle(Title);
		row = 2;
		sht2.createRow(row).createCell(col).setCellValue("Debtors as on " + reportdt);
		sht2.getRow(row).getCell(col).setCellStyle(Title);
		row = 4;
		col = 6;
		sht2.createRow(row);
		for (int a = 6; a < buckethead.length + 6; a++) {
			sht2.getRow(row).createCell(col).setCellValue(buckethead[a - 6][1]);
			sht2.getRow(row).getCell(col).setCellStyle(rightAligned);
			sht2.setColumnWidth(col++, 2900);
		}
		sht2.getRow(row).createCell(col).setCellValue("Credit");
		sht2.getRow(row).getCell(col).setCellStyle(rightAligned);
		sht2.setColumnWidth(col++, 2900);
		sht2.getRow(row).createCell(col).setCellValue("Total");
		sht2.getRow(row).getCell(col).setCellStyle(rightAligned);
		sht2.setColumnWidth(col++, 2900);
		sht2.getRow(row).createCell(col).setCellValue("Deposit");
		sht2.getRow(row).getCell(col).setCellStyle(rightAligned);
		sht2.setColumnWidth(col++, 2900);
		sht2.getRow(row).createCell(col).setCellValue("ECL Provn");
		sht2.getRow(row).getCell(col).setCellStyle(rightAligned);
		sht2.setColumnWidth(col++, 2900);
		col = 1;
		sht2.getRow(row).createCell(col).setCellValue("Code");
		sht2.getRow(row).getCell(col).setCellStyle(header);
		sht2.setColumnWidth(col++, 1800);
		sht2.getRow(row).createCell(col).setCellValue("Customer");
		sht2.getRow(row).getCell(col).setCellStyle(header);
		sht2.setColumnWidth(col++, 9000);
		sht2.getRow(row).createCell(col).setCellValue("Status");
		sht2.getRow(row).getCell(col).setCellStyle(header);
		sht2.setColumnWidth(col++, 2600);
		sht2.getRow(row).createCell(col).setCellValue("CustGroup");
		sht2.getRow(row).getCell(col).setCellStyle(header);
		sht2.setColumnWidth(col++, 3200);
		sht2.getRow(row).createCell(col).setCellValue("Pan No.");
		sht2.getRow(row).getCell(col).setCellStyle(header);
		sht2.setColumnWidth(col++, 3100);
		sht2.setColumnWidth(0, 400);
		sht2.setDisplayGridlines(false);
		// vertical total
		sl = sht2.getLastRowNum();
		sht2.createRow(sl + 2);
		for (col = 6; col < 17; col++) {
			// System.out.println(col);
			sht2.getRow(sl + 2).createCell(col)
					.setCellFormula("sum(" + (char) (97 + col) + "6:" + (char) (97 + col) + (sl + 1) + ")");
			sht2.getRow(sl + 2).getCell(col).setCellStyle(total);
		}
		sl = sht2.getLastRowNum();
		sht2.getRow(sl).createCell(5).setCellValue("Total");
		sht2.getRow(sl).getCell(5).setCellStyle(header2);
		sl = sht2.getLastRowNum();
		pt = new PropertyTemplate();
		pt.drawBorders(new CellRangeAddress(4, sl, 1, sht2.getRow(4).getLastCellNum() - 1), BorderStyle.THIN,
				IndexedColors.GREY_25_PERCENT.getIndex(), BorderExtent.ALL);
		pt.applyBorders(sht2);
		// file closing
		FileOutputStream f = new FileOutputStream(filepath + co[0] + "-Drs-" + reportdt + ".xlsx");
		wb.write(f);
		wb.close();
		Date end = new Date();
		long diff = end.getTime() - start.getTime();
		String TimeTaken = String.format("[%s] hours : [%s] mins : [%s] secs",
				Long.toString(TimeUnit.MILLISECONDS.toHours(diff)), TimeUnit.MILLISECONDS.toMinutes(diff),
				TimeUnit.MILLISECONDS.toSeconds(diff));
		System.out.println(String.format("Completed ...... Time taken %s", TimeTaken));
	}
}
