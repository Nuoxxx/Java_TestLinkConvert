package com.sky.testlinkconvert;

import java.io.*;
import java.text.DecimalFormat;
import java.util.*;

import org.apache.poi.POIXMLException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.*;
import org.dom4j.io.*;

/**
 * 将写有用例的excel文件转成xml文件,均以套件形式导入，支持2类模板: 模板一（固定4列）： [测试序号] 用例名称 预置条件 操作步骤 预期结果 +
 * [随机额外列（含用例等级）] 模板二（固定6列）： [测试序号] 模块 子模块 用例名称 预置条件 操作步骤 预期结果 + [随机额外列（含用例等级）]
 * 注意： 以上固定的列，要求必须有，且顺序必须跟模板相同，
 * []括起的列可以有也可以没有，如果有的话，"测试序号"和"用例等级"名称必须如此，其他额外列名称可自定义；
 * 测试序号若有，必须位于第一列，其他额外列放在"预期结果"后，顺序随意；
 * []括起的列，除了用例等级，其他的信息都将导入到testlink中用例的"摘要"信息中。
 * 
 *
 */
public class ExcelToXml {
	private static int internalid = 1000001;
	private static int suite_node = 1;
	private static List<String> s_modules = new ArrayList<String>(); // 保存：子模块名，重复只保存一次
	// excel模板支持4个/6个固定列，及多个随机列；extracols不为empty，信息计入摘要中
	private static List<String> titles = new ArrayList<String>(); // 放所有列名称
	private static List<String> extracols = new ArrayList<String>(); // 放额外的标题名称
	private static int yq_index; // 预期结果列的下标
	private static int m_index; // 模块列的下标
	private static int sm_index; // 子模块列的下标
	private static int totalCol;
	private static String modulesString;

	// private static String module;

	public static void transferExcelToXml(String oldfilename) {
		long time;
		String newfilename;
		// System.out.println("newfilename:"+newfilename);
		System.out.println("converting,please wait...");

		// 初始化static属性值
		internalid = 1000001;
		suite_node = 1;
		s_modules.clear();
		titles.clear();
		extracols.clear();

		// 默认创建2007版本的Excel文件对象
		XSSFWorkbook xswb = null;
		// 出现异常时，创建2003版本的Excel文件对象
		HSSFWorkbook hswb = null;
		try {
			//
			System.out.println("oldfilename:" + oldfilename);
			xswb = new XSSFWorkbook(new FileInputStream(oldfilename));
			// 创建对工作表的引用
			XSSFSheet xssheet;

			for (int sheetNum = 0; sheetNum < xswb.getNumberOfSheets(); sheetNum++) {
				// 清空很重要，不然后面的页签会出现多个摘要
				s_modules.clear();
				titles.clear();
				extracols.clear();
				//得到Excel工作表对象
				xssheet = xswb.getSheetAt(sheetNum);
				time = System.currentTimeMillis();
				//获取工作表的名字作为模块名
				modulesString = xssheet.getSheetName();
				System.out.println("表格名：" + modulesString);
				//用模块名作为生成的XML文件名
				newfilename = getXmlName(oldfilename, time, modulesString);
				

				List<String> caseatrs;
				List<String> sm_caseatrs = new ArrayList<String>();
				String tempfile = ""; // 临时文件的定义

				// 获取行号
				int totalRow = xssheet.getLastRowNum();
				System.out.println("totalRow91:" + totalRow);

				// 获取标题行
				XSSFRow xsrow0 = xssheet.getRow(0);
				// 获取标题行的列数
				totalCol = xsrow0.getLastCellNum();
				// 将列标题保存起来
				for (int r = 0; r < totalCol; r++) {
					titles.add(xsrow0.getCell(r).getStringCellValue());
				}
				yq_index = titles.indexOf("预期结果");

				// //删除文档最后的空行
				for (int i = totalRow; i > 1; i--) {
					System.out.println("i:" + i);

					XSSFCell xCell = xssheet.getRow(i).getCell(yq_index);
					// System.out.println("第"+i+"行数据："+xCell);

					if (xCell == null) {
						totalRow--;
						System.out.println("totalRow:" + totalRow);
					} else {
						String xData = xssheet.getRow(i).getCell(yq_index).getStringCellValue();
						if (xData.replaceAll("\\s", "").equals("")) {
							totalRow--;
							System.out.println("有空白，totalRow:" + totalRow);
						} else {
							System.out.println("有数据了，停止循环");
							break;
						}
					}
				}

				System.out.println("删除文档最后空行后总行数：" + totalRow);
				// 若预期结果不是最后一列，则有额外列，保存额外列名称
				for (int cn = yq_index + 1; cn < totalCol; cn++) {
					extracols.add(xsrow0.getCell(cn).getStringCellValue());
				}

				// 若存在模块列和子模块列，记录下这两列的下标
				if (yq_index - 4 >= 0 && yq_index - 5 >= 0) {
					// m_index = yq_index-5;
					sm_index = yq_index - 4;
				}
				sm_caseatrs.add("");
				// 处理子模块数据
				for (int i = 1; i <= totalRow; i++) {
					// 增加一个空行，为了和Excel中的title行对应
					XSSFCell s_module = xssheet.getRow(i).getCell(0);
					System.out.println("s_module:" + s_module);
					if (s_module != null && !s_module.equals("")) {
						if (isMergedRegion(xssheet, s_module)) {

							int moduleMergeRow = mergedRow(xssheet, s_module);
//							String moduleNameString = getCellContent(s_module);
							for (int k = 0; k < moduleMergeRow + 1; k++) {
//								sm_caseatrs.add(moduleNameString);
								sm_caseatrs.add(replaceCellAngleBrackets(s_module.getStringCellValue()));
								System.out.println("为合并单元格设置数据");
							}
							i = i + moduleMergeRow;
						} else {
							sm_caseatrs.add(replaceCellAngleBrackets(s_module.getStringCellValue()));
						}
					} else {
						sm_caseatrs.add("");
					}
				}
				// 打印出所有数据
				// for (int i = 0; i < sm_caseatrs.size(); i++) {
				// System.out.println("sm_caseatrs中的数据："+sm_caseatrs.get(i));
				// }
				Iterator<String> iterator = sm_caseatrs.iterator();
				while (iterator.hasNext()) {
					System.out.println("sm_caseatrs中的数据：" + iterator.next());
				}
				// 获取用例各行信息
				for (int row = 1; row <= totalRow; row++) {

					caseatrs = new ArrayList<String>();
					caseatrs.add(sm_caseatrs.get(row));

					for (int col = 1; col < totalCol; col++) {
						XSSFCell temp = xssheet.getRow(row).getCell(col);
						System.out.println("XSSFCell temp:"+temp);
						// 判断是否是用例名合并单元格

						if (temp != null && !temp.equals("") && isMergedRegion(xssheet, temp)) {
							// 自己写的，获取合并单元格的总数
							System.out.println("合并单元格的数量:" + xssheet.getNumMergedRegions());
							// 获取合并单元格行数
							int mergedRowNum = mergedRow(xssheet, temp);
							System.out.println("mergedRowNum:"+mergedRowNum);
							caseatrs.clear();

							for (int r = 0; r < mergedRowNum + 1; r++) {
								caseatrs.add(sm_caseatrs.get(row));
								for (int k = 1; k < totalCol; k++) {
									XSSFCell mergedTempCell = xssheet.getRow(row + r).getCell(k);
									System.out.println("mergedTempCell:"+mergedTempCell);
									// 将数据增加到caseatrs中

									if (mergedTempCell != null && !mergedTempCell.equals("")) {
										if (mergedTempCell.getCellType() == mergedTempCell.CELL_TYPE_NUMERIC) {
											DecimalFormat df = new DecimalFormat("###0");
											caseatrs.add(replaceCellAngleBrackets(
													df.format(mergedTempCell.getNumericCellValue())));
										} else {
											caseatrs.add(replaceCellAngleBrackets(mergedTempCell.getStringCellValue()));
										}
									} else {
										caseatrs.add("");
									}
								}
							}
							row = row + mergedRowNum;
							break;
						} else {
							System.out.println("不是用例名合并单元格");
							if (temp != null && !temp.equals("")) {
								System.out.println("temp的值:" + temp.toString());
								if (temp.getCellType() == temp.CELL_TYPE_NUMERIC) {
									DecimalFormat df = new DecimalFormat("###0");
									caseatrs.add(replaceCellAngleBrackets(df.format(temp.getNumericCellValue())));
								} else {
									caseatrs.add(replaceCellAngleBrackets(temp.getStringCellValue()));
								}

							} else {
								System.out.println("加空数据");
								caseatrs.add("");
							}
						}
					}

					// 打印出所有数据
					for (int i = 0; i < caseatrs.size(); i++) {
						System.out.println("caseatrs中的数据：" + caseatrs.get(i));
					}

					// 先写到临时xml文件中,创建临时文件所在目录，再完成临时文件的赋值
					File temp = new File("c:");
					if (temp.exists()) {// 有c盘
						if (!new File("c:\\temp").exists()) {
							new File("c:\\temp").mkdirs();
						}
						tempfile = "c:\\temp\\tempfile_" + time + ".xml";
					} else {// 无c盘
						String path = oldfilename.split("\\\\")[0];
						if (!new File(path + "\\temp").exists()) {
							new File(path + "\\temp").mkdirs();
						}
						tempfile = path + "\\temp\\tempfile_" + time + ".xml";
					}

					int result = writeTestcaseToXml(tempfile, caseatrs);
					if (result == -1) {
						return;
					}

				}
				// 读完一张表中内容，将临时文件的内容重新写到最终的xml文件中
				replaceESC(tempfile, newfilename);

			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (POIXMLException e) {
			try {
				// Excel2003
				System.out.println("Excel2003");
				hswb = new HSSFWorkbook(new FileInputStream(oldfilename));
				// HSSFSheet hssheet = hswb.getSheetAt(0);
				HSSFSheet hssheet;
				// System.out.println("sheetNum:"+);
				// 不同的表格对应不同的xml文件
				for (int sheetNum = 0; sheetNum < hswb.getNumberOfSheets(); sheetNum++) {

					s_modules.clear();
					titles.clear();
					extracols.clear();

					hssheet = hswb.getSheetAt(sheetNum);
					time = System.currentTimeMillis();

					modulesString = hssheet.getSheetName();
					newfilename = getXmlName(oldfilename, time, modulesString);
					System.out.println("表格名" + modulesString);

					List<String> caseatrs;
					List<String> sm_caseatrs = new ArrayList<String>();
					String tempfile = ""; // 临时文件的定义

					// 获取行号
					int totalRow = hssheet.getLastRowNum();
					// 获取标题行的列数
					HSSFRow hsrow0 = hssheet.getRow(0);
					totalCol = hsrow0.getLastCellNum();
					// 将列标题保存起来
					for (int r = 0; r < totalCol; r++) {
						titles.add(hsrow0.getCell(r).getStringCellValue());
					}
					yq_index = titles.indexOf("预期结果");
					// 若预期结果不是最后一列，则有额外列，保存额外列名称
					for (int cn = yq_index + 1; cn < totalCol; cn++) {
						extracols.add(hsrow0.getCell(cn).getStringCellValue());
					}

					// 若存在模块列和子模块列，记录下这两列的下标
					if (yq_index - 4 >= 0 && yq_index - 5 >= 0) {
						// m_index = yq_index-5;
						sm_index = yq_index - 4;
						System.out.println("yq_index:" + yq_index + "sm_index:" + sm_index);
					}
					sm_caseatrs.add("");
					// 处理子模块数据
					for (int i = 1; i <= totalRow; i++) {
						// 增加一个空行，为了和Excel中的title行对应
						HSSFCell moduleSet = hssheet.getRow(i).getCell(0);
						if (moduleSet != null && !moduleSet.equals("")) {
							if (isMergedRegion(hssheet, moduleSet)) {

								int moduleMergeRow = mergedRow(hssheet, moduleSet);
								String moduleNameString = getCellContent(moduleSet);
								for (int k = 0; k < moduleMergeRow + 1; k++) {
									sm_caseatrs.add(moduleNameString);
									// System.out.println("为合并单元格设置数据");
								}
								i = i + moduleMergeRow;
							} else {
								sm_caseatrs.add(replaceCellAngleBrackets(moduleSet.getStringCellValue()));
							}
						} else {
							sm_caseatrs.add("");
						}
					}
					// 打印出所有数据
					// for (int i = 0; i < sm_caseatrs.size(); i++) {
					// System.out.println("sm_caseatrs中的数据："+sm_caseatrs.get(i));
					// }
					Iterator<String> iterator = sm_caseatrs.iterator();
					while (iterator.hasNext()) {
						System.out.println("sm_caseatrs中的数据：" + iterator.next());
					}
					// 获取用例的各列信息
					for (int row = 1; row <= totalRow; row++) {
						caseatrs = new ArrayList<String>();
						caseatrs.add(sm_caseatrs.get(row));

						for (int col = 1; col < totalCol; col++) {
							HSSFCell temp = hssheet.getRow(row).getCell(col);

							if (temp != null && !temp.equals("") && isMergedRegion(hssheet, temp)) {
								// System.out.println("合并单元格的数量:"+hssheet.getNumMergedRegions());
								int mergedRow = mergedRow(hssheet, temp);
								caseatrs.clear();

								for (int r = 0; r < mergedRow + 1; r++) {
									caseatrs.add(sm_caseatrs.get(row));
									for (int k = 1; k < totalCol; k++) {
										HSSFCell mergedTempCell = hssheet.getRow(row + r).getCell(k);
										// 将数据增加到caseatrs中

										if (mergedTempCell != null && !mergedTempCell.equals("")) {
											if (mergedTempCell.getCellType() == mergedTempCell.CELL_TYPE_NUMERIC) {
												DecimalFormat df = new DecimalFormat("###0");
												caseatrs.add(replaceCellAngleBrackets(
														df.format(mergedTempCell.getNumericCellValue())));
											} else {
												caseatrs.add(
														replaceCellAngleBrackets(mergedTempCell.getStringCellValue()));
											}
										} else {
											caseatrs.add("");
										}
									}
								}
								row = row + mergedRow;
								break;
							} else {
								System.out.println("不是用例名合并单元格");
								if (temp != null && !temp.equals("")) {
									System.out.println("temp的值:" + temp.toString());
									if (temp.getCellType() == temp.CELL_TYPE_NUMERIC) {
										DecimalFormat df = new DecimalFormat("###0");
										caseatrs.add(replaceCellAngleBrackets(df.format(temp.getNumericCellValue())));
									} else {
										caseatrs.add(replaceCellAngleBrackets(temp.getStringCellValue()));
									}

								} else {
									// System.out.println("加空数据");
									caseatrs.add("");
								}
							}
						}

						// 打印出所有数据
						for (int i = 0; i < caseatrs.size(); i++) {
							System.out.println("caseatrs中的数据：" + caseatrs.get(i));
						}

						// 先写到临时xml文件中,创建临时文件所在目录，再完成临时文件的赋值
						File temp = new File("c:");
						if (temp.exists()) {// 有c盘
							if (!new File("c:\\temp").exists()) {
								new File("c:\\temp").mkdirs();
							}
							tempfile = "c:\\temp\\tempfile_" + time + ".xml";
						} else {// 无c盘
							String path = oldfilename.split("\\\\")[0];
							if (!new File(path + "\\temp").exists()) {
								new File(path + "\\temp").mkdirs();
							}
							tempfile = path + "\\temp\\tempfile_" + time + ".xml";
						}

						int result = writeTestcaseToXml(tempfile, caseatrs);
						if (result == -1) {
							return;
						}
					}
					// 将临时文件的内容重新写到最终的xml文件中
					replaceESC(tempfile, newfilename);
				}

			} catch (FileNotFoundException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (NullPointerException e1) {
				// TODO Auto-generated catch block
				System.out.println("In 2003 excel,TestCase and TestCase between can't have empty row.");
				System.out.println("OR some columns or some rows are not in border.");
				e1.printStackTrace();
			} catch (Exception e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
		} catch (IOException e) {
			e.printStackTrace();
		} catch (NullPointerException e1) {
			// TODO Auto-generated catch block
			System.out.println(e1.getStackTrace());
			System.out.println("In 2007 excel,TestCase and TestCase between can't have empty row.");
			System.out.println("OR some columns or some rows are not in border.");
			e1.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * 一行一行的将每个测试用例写入xml中
	 * 
	 * 根据模块和子模块名称，添加具体的测试套件，模块和子模块各对应一个 模块和子模块为空时，不添加对应的测试套件
	 */
	private static int writeTestcaseToXml(String newfilename, List<String> caseatrs) {
		System.out.println("开始写xml");
		try {
			XMLWriter writer = null;// 声明写XML的对象
			//SAXReader(Stream API for XML)把XML文档作为一个流来处理
			SAXReader reader = new SAXReader();
			Document document = null;

			OutputFormat format = OutputFormat.createPrettyPrint();
			// 由之前的UTF-8编码改成gb2312编码格式，之前的UTF-8编码，从Jar包生成xml时文件的编码格式会出错
			format.setEncoding("UTF-8");// 设置XML文件的编码格式

			File file = new File(newfilename);
			/*判断文件是否存在
			 * 1:存在，追加测试用例
			 * 2.不存在，新建文件并追加测试用例
			 */
			if (file.exists()) { 
				// 读取存在的testcase.xml文件，并追加测试用例
				// document = reader.read(file); // 读取XML文件
				// 之前直接生成Jar包会出故障
				document = reader.read(new BufferedReader(new InputStreamReader(new FileInputStream(file))));
				Element root = document.getRootElement(); // 得到根节点

				// 无模块和子模块列，直接将用例添加到根节点下
				if (yq_index - 4 < 0) {
					System.out.println("无模块和子模块");
					addTestCase(root, caseatrs);

				} else {// 有子模块列
						// 子模块列有值
						// 排除空数据
					if (caseatrs.get(yq_index) != null && !caseatrs.get(yq_index).equals("")) {
						if (!caseatrs.get(sm_index).equals("")) {
							String s_module = caseatrs.get(sm_index);
							System.out.println("子模块值:" + s_module);
							// 子模块名都已有对应的测试套件
							if (s_modules.contains(s_module)) {
								System.out.println("已经包含该子模块");
								// 获取已有测试套件，并添加测试用例
								Element element = getTestsuiteByModule(root, s_module);
								// 添加测试用例
								// 打印modules中的值
								for (int m = 0; m < s_modules.size(); m++) {
									System.out.println("modules中的值：" + s_modules.get(m));
								}

								System.out.println("已包含模块，直接添加用例" + element.toString());
								addTestCase(element, caseatrs);
							} else { // 子模块中没有对应的测试套件
								System.out.println("添加新子模块");
								Element sub_testsuite = root.addElement("testsuite");
								sub_testsuite.addAttribute("name", caseatrs.get(sm_index));
								Element sub_node = sub_testsuite.addElement("node_order");
								sub_node.setText("<![CDATA[" + suite_node + "]]>");
								Element sub_details = sub_testsuite.addElement("details");
								sub_details.setText("<![CDATA[]]>");
								suite_node++;
								// 在子测试套件下添加测试用例
								addTestCase(sub_testsuite, caseatrs);
								// 保存新建模块
								s_modules.add(s_module);
							}

						} else if (caseatrs.get(sm_index).equals("")) {
							// 子模块列无值，不建测试套件，直接添加测试用例
							System.out.println("直接添加testcase");
							addTestCase(root, caseatrs);

						} else {
							System.out.println(
									"converting Fail! Caused by:module name is empty when child module name is not empty!");
							return -1;
						}
					}
				}
			} else {
				// 新建testcase.xml文件
				document = DocumentHelper.createDocument();
				// 建根节点
				Element root = document.addElement("testsuite");
				root.addAttribute("name", modulesString);

				// 无模块和子模块列，直接将用例添加到根节点下
				if (yq_index - 4 < 0) {
					addTestCase(root, caseatrs);

				} else {
					// 有子模块列
					/* 子模块是否有值：
					 * 1.有值，子模块列写入XML
					 * 2.无值，直接添加测试用例
					 */					
					if (!caseatrs.get(sm_index).equals("")) {
						System.out.println("有子模块列");
						//获取子模块名,添加子模块信息  eg.<testsuite name="直播信息xml"> 
						String s_module = caseatrs.get(sm_index);
						Element sub_testsuite = root.addElement("testsuite");
						sub_testsuite.addAttribute("name", caseatrs.get(sm_index));
						//
						Element sub_node = sub_testsuite.addElement("node_order");
						//suite_node对应？？
						sub_node.setText("<![CDATA[" + suite_node + "]]>");
						Element sub_details = sub_testsuite.addElement("details");
						sub_details.setText("<![CDATA[]]>");
						suite_node++;
						
						// 在子测试套件下添加测试用例
						addTestCase(sub_testsuite, caseatrs);
						// 保存新建模块
						s_modules.add(s_module);
					} else if (caseatrs.get(sm_index).equals("")) {
						// 子模块无值，不建测试套件，直接添加测试用例
						addTestCase(root, caseatrs);

					} else {
						System.out.println(
								"converting Fail! Caused by:module name is empty when child module name is not empty!");
						return -1;
					}
				}
			}
			// writer = new XMLWriter(new FileWriter(newfilename), format);
			// 之前直接生成Jar包会出故障
			writer = new XMLWriter(new OutputStreamWriter(new FileOutputStream(new File(newfilename))), format);

			writer.write(document);
			writer.close();
			internalid++;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return 0;
	}

	// 根据模块名创建外层测试套件
	private static Element createTestsuitesByModule(Element root, List<String> caseatrs) {
		// 建外层测试套件
		Element child_element = null;
		if (!caseatrs.get(m_index).contains("/")) {
			Element testsuite = root.addElement("testsuite");
			testsuite.addAttribute("name", caseatrs.get(m_index));
			Element node = testsuite.addElement("node_order");
			node.setText("<![CDATA[" + suite_node + "]]>");
			Element details = testsuite.addElement("details");
			details.setText("<![CDATA[]]>");
			suite_node++;
			child_element = testsuite;
		} else {
			String[] suite_names = caseatrs.get(m_index).split("/");
			Element temp = root;
			for (String suite_name : suite_names) {
				Element testsuite = temp.addElement("testsuite");
				testsuite.addAttribute("name", suite_name);
				Element node = testsuite.addElement("node_order");
				node.setText("<![CDATA[" + suite_node + "]]>");
				Element details = testsuite.addElement("details");
				details.setText("<![CDATA[]]>");
				suite_node++;
				temp = testsuite;
			}
			child_element = temp;
		}

		return child_element;
	}

	//返回子模块对应的testsuite
	private static Element getTestsuiteByModule(Element root, String module) {
		Element testSuite = null;
		//遍历根目录下所有testsuite（子模块）
		Iterator<Element> it = root.elementIterator("testsuite");

		while (it.hasNext()) {
			Element element = it.next();
			if (element.attributeValue("name").equals(module)) {
				testSuite = element;
			}
		}
		return testSuite;
	}

	private static void addTestCase(Element sup_element, List<String> caseatrs) {
		// 多个操作步骤

		// 添加一个testcase
		Element testcase = sup_element.addElement("testcase");
		testcase.addAttribute("internalid", internalid + "");
		testcase.addAttribute("name", caseatrs.get(yq_index - 3));

		// 将测试序号及额外列（除开用例等级）信息导入到summary（摘要）中
		Element summary = testcase.addElement("summary");
		StringBuffer sumStr = new StringBuffer();
		if (titles.indexOf("测试序号") == 0) {
			sumStr.append("测试序号：" + caseatrs.get(0));
		}

		for (String col : extracols) {
			if (col.equals("摘要")) {
				System.out.println("col:" + col);
				sumStr.append("</br>");
				sumStr.append(caseatrs.get(extracols.indexOf(col) + yq_index + 1).replaceAll("\n", "</br>"));
				System.out.println("摘要append:" + caseatrs.get(extracols.indexOf(col) + yq_index + 1).replaceAll("\n", "</br>"));
			}
			// else if(!col.equals("用例等级")){
			// sumStr.append("</br>");
			// sumStr.append(col+"："+caseatrs.get(extracols.indexOf(col)+yq_index+1));
			// }
		}
		System.out.println("sumStr:"+sumStr.toString());
		summary.setText("<![CDATA[" + sumStr.toString() + "]]>");

		Element preconditions = testcase.addElement("preconditions");
		preconditions.setText("<![CDATA[" + caseatrs.get(yq_index - 2).replaceAll("\n", "</br>") + "]]>");
		//执行方式：0：自动化测试；1：手工测试；
		Element execution_type = testcase.addElement("execution_type");
		execution_type.setText("<![CDATA[1]]>");

		// 额外列中，如果有用例等级，取对应的用例等级导入；如果无用例等级，默认用例等级为"2"
		Element importance = testcase.addElement("importance");
		int index = extracols.indexOf("用例等级");
		if (index != -1) {
			index = (yq_index + 1) + index;
			if (caseatrs.get(index).equals("低") || caseatrs.get(index).equals("1")) {
				importance.setText("<![CDATA[" + 1 + "]]>");
			} else if (caseatrs.get(index).equals("中") || caseatrs.get(index).equals("2")) {
				importance.setText("<![CDATA[" + 2 + "]]>");
			} else if (caseatrs.get(index).equals("高") || caseatrs.get(index).equals("3")) {
				importance.setText("<![CDATA[" + 3 + "]]>");
			}
		} else {
			importance.setText("<![CDATA[" + 2 + "]]>");
		}

		Element steps = testcase.addElement("steps");
		System.out.println("caseatrs.size()：" + caseatrs.size());
		//通过list长度和表格的列相除，对应添加多步骤
		int stepNumber = caseatrs.size() / totalCol;
		System.out.println("stepNumber：" + stepNumber);
		for (int i = 0; i < stepNumber; i++) {
			Element step = steps.addElement("step");
			//第几步
			Element step_number = step.addElement("step_number");
			step_number.setText("<![CDATA[" + (i + 1) + "]]>");
			//操作步骤
			Element actions = step.addElement("actions");
			actions.setText(
					"<![CDATA[" + caseatrs.get((yq_index - 1) + totalCol * i).replaceAll("\n", "</br>") + "]]>");
			//操作结果
			Element expectedresults = step.addElement("expectedresults");
			expectedresults
					.setText("<![CDATA[" + caseatrs.get(yq_index + totalCol * i).replaceAll("\n", "</br>") + "]]>");
		}

		// 自己添加的定义区域
		Element custom_fields = testcase.addElement("custom_fields");
		Element custom_field = custom_fields.addElement("custom_field");

		Element name = custom_field.addElement("name");
		name.setText("<![CDATA[author]]>");
		Element value = custom_field.addElement("value");
		value.setText("<![CDATA[" + caseatrs.get(yq_index + 2) + "]]>");
	}

	// 根据老的文件名 获取新的文件名
	private static String getXmlName(String oldfilename, long time, String SheetName) {
		String newfilename = "";
		String[] temp = oldfilename.split("\\\\");
		String name = temp[temp.length - 1].split("\\.")[0]; // 文件名前缀
		name = name.replaceAll("[0-9]*", "");
		if (name.endsWith("_") == true) {
			newfilename = oldfilename.substring(0, oldfilename.length() - temp[temp.length - 1].length()) + "TestCase_"
					+ SheetName + "_" + time + ".xml";
		} else {
			newfilename = oldfilename.substring(0, oldfilename.length() - temp[temp.length - 1].length()) + "TestCase_"
					+ SheetName + "_" + time + ".xml";
		}
		return newfilename;
	}

	// 简单替换每列内容中的<>符号为小于和大于；最好要求用户不要使用尖括号，否则会替换成大于小于
	private static String replaceCellAngleBrackets(String cellStr) throws Exception {
		String result = "";
		if (cellStr.contains("<") && cellStr.contains(">")) {
			result = cellStr.replaceAll("<", "小于");
			result = result.replaceAll(">", "大于");
		} else if (cellStr.contains("<")) {
			result = cellStr.replaceAll("<", "小于");
		} else if (cellStr.contains(">")) {
			result = cellStr.replaceAll(">", "大于");
		} else {
			result = cellStr;
		}
		return result;
	}

	// 替换xml文件中的转义字符
	private static void replaceESC(String tempfile, String newfilename) throws Exception {
		File file = new File(tempfile);
		FileInputStream fis = new FileInputStream(tempfile);
		InputStreamReader isr = new InputStreamReader(fis);
		BufferedReader br = new BufferedReader(isr);

		FileOutputStream fos = new FileOutputStream(newfilename, true);
		OutputStreamWriter osw = new OutputStreamWriter(fos);
		BufferedWriter bw = new BufferedWriter(osw);

		// 一行一行的读，一行一行的写
		String line;
		while ((line = br.readLine()) != null) {
			String tempstr = line.replaceAll("&lt;", "<");
			tempstr = tempstr.replaceAll("&gt;", ">");
			bw.write((tempstr + "\n"));
		}
		br.close();
		isr.close();
		fis.close();
		bw.close();
		osw.close();
		fos.close();

		// 删除临时文件
		file.delete();
	}

	// 判断是否是单元格
	public static boolean isMergedRegion(Sheet sheet, Cell cell) {

		// 得到一个sheet中有多少个合并单元格
		int sheetmergerCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetmergerCount; i++) {
			// 得出具体的合并单元格
			CellRangeAddress ca = sheet.getMergedRegion(i);
			// 得到合并单元格的起始行, 结束行, 起始列, 结束列
			int firstC = ca.getFirstColumn();
			int lastC = ca.getLastColumn();
			int firstR = ca.getFirstRow();
			int lastR = ca.getLastRow();
			// 判断该单元格是否在合并单元格范围之内, 如果是, 则返回 true
			if (cell.getColumnIndex() <= lastC && cell.getColumnIndex() >= firstC) {
				if (cell.getRowIndex() <= lastR && cell.getRowIndex() >= firstR) {
					System.out.println("lastR-firstR:" + (lastR - firstR));
					return true;
				}
			}
		}
		return false;
	}

	// 获取合并单元格的行数
	public static int mergedRow(Sheet sheet, Cell cell) {
		int sheetmergerCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetmergerCount; i++) {
			// 得出具体的合并单元格
			CellRangeAddress ca = sheet.getMergedRegion(i);
			// 得到合并单元格的起始行, 结束行, 起始列, 结束列
			int firstC = ca.getFirstColumn();
			int lastC = ca.getLastColumn();
			int firstR = ca.getFirstRow();
			int lastR = ca.getLastRow();
			// 判断该单元格是否在合并单元格范围之内, 如果是, 则返回 true
			if (cell.getColumnIndex() <= lastC && cell.getColumnIndex() >= firstC) {
				if (cell.getRowIndex() <= lastR && cell.getRowIndex() >= firstR) {
					System.out.println("合并单元格的行数:" + (lastR - firstR));
					return lastR - firstR;
				}
			}
		}
		return 0;
	}

	public static String getMergedRegionValue(Sheet sheet, Cell cell) {
		// 获得一个 sheet 中合并单元格的数量
		int sheetmergerCount = sheet.getNumMergedRegions();
		// 便利合并单元格
		for (int i = 0; i < sheetmergerCount; i++) {
			// 获得合并单元格
			CellRangeAddress ca = sheet.getMergedRegion(i);
			// 获得合并单元格的起始行, 结束行
			int firstC = ca.getFirstColumn();
			int lastC = ca.getLastColumn();
			int firstR = ca.getFirstRow();
			int lastR = ca.getLastRow();
			if (cell.getColumnIndex() == firstC && cell.getRowIndex() == firstR) {
				return getCellContent(cell);

			}

		}
		return "";

	}

	private static final DataFormatter FORMATTER = new DataFormatter();

	private static String getCellContent(Cell cell) {
		return FORMATTER.formatCellValue(cell);
	}

}
