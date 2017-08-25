package com.sky.testlinkconvert;

import java.awt.Color;
import java.awt.FileDialog;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetDragEvent;
import java.awt.dnd.DropTargetDropEvent;
import java.awt.dnd.DropTargetEvent;
import java.awt.dnd.DropTargetListener;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;

/**
 * testlinkconvert的图形界面 Alt+shift+r 批量修改变量名
 */
public class ConvertGui implements ActionListener {

	JTextField jtext = new DropDragSupportTextArea(20);
	FileDialog fd = null;
	JFrame jf = null;
	JButton jb_ChooseFile = null;
	JButton jb_XmlToExcel = null;
	JButton jb_ExcelToXml = null;
	String oldfilename;
	boolean isExcelToXml = false;

	// 转换器界面布局
	public ConvertGui() {
		jf = new JFrame("Testlink转换器");
		fd = new FileDialog(jf);
		JPanel jp_file = new JPanel();
		JPanel jp_transfer_ype = new JPanel();
		JLabel jlable = new JLabel("源文件:");
		// 三个按钮
		jb_ChooseFile = new JButton("选择");
		jb_XmlToExcel = new JButton("xml转成excel");
		jb_ExcelToXml = new JButton("excel转成xml");
		jb_ChooseFile.addActionListener(this);
		jb_XmlToExcel.addActionListener(this);
		jb_ExcelToXml.addActionListener(this);
		// 上层JPanel 增加 “源文件”、“文件显示框”、“选择按钮”
		jp_file.add(jlable);
		jtext.setEditable(false);
		jtext.setBackground(Color.white);
		jtext.addActionListener(this);
		jp_file.add(jtext);
		jp_file.add(jb_ChooseFile);
		// 下层JPanel 增加“xml转成Excel”、“Excel转成XML”按钮
		jb_XmlToExcel.setEnabled(false);
		jb_ExcelToXml.setEnabled(false);
		jp_transfer_ype.add(jb_XmlToExcel);
		jp_transfer_ype.add(jb_ExcelToXml);

		// 将以上两个JPanel加入JFrame中
		jf.add(jp_file, "North");
		jf.add(jp_transfer_ype);
		jf.setLocation(300, 200);
		jf.setVisible(true);
		jf.pack();
		jf.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	}

	public static void main(String[] args) {
		new ConvertGui();

	}

	public void actionPerformed(ActionEvent e) {
		String comm = e.getActionCommand();
		if (comm.equals("选择")) {
			fd.setVisible(true);
			if (fd.getFile() != null) {
				if (fd.getFile().endsWith(".xml") || fd.getFile().endsWith(".xls") || fd.getFile().endsWith(".xlsx")) {
					jtext.setText(fd.getDirectory() + fd.getFile());
					oldfilename = fd.getDirectory() + fd.getFile();

					if (fd.getFile().endsWith(".xml")) {
						jb_XmlToExcel.setEnabled(true);
						jb_ExcelToXml.setEnabled(false);
						isExcelToXml = false;
						jb_XmlToExcel.setEnabled(false);
						jb_ChooseFile.setEnabled(false);
						new FileTransferTool(jtext, jb_ChooseFile, jb_XmlToExcel, jb_ExcelToXml, oldfilename,
								isExcelToXml).start();
					} else {
						jb_ExcelToXml.setEnabled(true);
						jb_XmlToExcel.setEnabled(false);
						isExcelToXml = true;
						jb_ExcelToXml.setEnabled(false);
						jb_ChooseFile.setEnabled(false);
						new FileTransferTool(jtext, jb_ChooseFile, jb_XmlToExcel, jb_ExcelToXml, oldfilename,
								isExcelToXml).start();
					}
				} else {
					System.out.print("" + System.getProperties().getProperty("os.name"));
					JOptionPane.showMessageDialog(jtext, "请重新选择xml或excel文件！");
					jtext.setText("");
					jb_XmlToExcel.setEnabled(false);
					jb_ExcelToXml.setEnabled(false);
				}
			}
		} else {
			System.out.println("拖动");
			new DropDragSupportTextArea(0);
		}
	}

	class DropDragSupportTextArea extends JTextField implements DropTargetListener {
		private DropTarget dropTarget;

		public DropDragSupportTextArea(int arg0) {
			super(arg0);
			// 注册DropTarget，并将它与组件相连，处理哪个组件的相连
			// 即连通组件（第一个this）和Listener(第二个this)
			dropTarget = new DropTarget(this, DnDConstants.ACTION_COPY_OR_MOVE, this, true);
		}

		/**
		 * 拖入文件或字符串,这里只说明能拖拽，并未打开文件并显示到文本区域中
		 */
		public void dragEnter(DropTargetDragEvent dtde) {
			// TODO Auto-generated method stub
			System.out.println("dragEnter");
			DataFlavor[] dataFlavors = dtde.getCurrentDataFlavors();
			if (dataFlavors[0].match(DataFlavor.javaFileListFlavor)) {
				try {
					Transferable tr = dtde.getTransferable();
					Object obj = tr.getTransferData(DataFlavor.javaFileListFlavor);
					List<File> files = (List<File>) obj;

					if (files != null && files.size() > 0) {
						String absolutePath = files.get(0).getAbsolutePath();
						jtext.setText(absolutePath);
					}
					/*
					 * for (int i = 0; i < files.size(); i++) {
					 * append(files.get(i).getAbsolutePath() + "/r/n"); }
					 */
				} catch (UnsupportedFlavorException ex) {

				} catch (IOException ex) {

				}
			}
		}

		public void dragExit(DropTargetEvent dte) {
			// TODO Auto-generated method stub
			System.out.println("dragExit");
		}

		public void dragOver(DropTargetDragEvent dtde) {
			// TODO Auto-generated method stub
			System.out.println("dragOver");
		}

		public void drop(DropTargetDropEvent dtde) {
			// TODO Auto-generated method stub

			System.out.println("drop");
			System.out.println(fd.getFile());
			if (jtext != null) {
				System.out.println("开始");
				if (jtext.getText().endsWith(".xml") || jtext.getText().endsWith(".xls")
						|| jtext.getText().endsWith(".xlsx")) {
					// ja1.setText(fd.getDirectory()+fd.getFile());
					// oldfilename=fd.getDirectory()+fd.getFile();
					oldfilename = jtext.getText();
					if (jtext.getText().endsWith(".xml")) {
						jb_XmlToExcel.setEnabled(true);
						jb_ExcelToXml.setEnabled(false);
						isExcelToXml = false;
						jb_XmlToExcel.setEnabled(false);
						jb_ChooseFile.setEnabled(false);
						new FileTransferTool(jtext, jb_ChooseFile, jb_XmlToExcel, jb_ExcelToXml, oldfilename,
								isExcelToXml).start();
					} else {
						jb_ExcelToXml.setEnabled(true);
						jb_XmlToExcel.setEnabled(false);
						isExcelToXml = true;
						jb_ExcelToXml.setEnabled(false);
						jb_ChooseFile.setEnabled(false);
						new FileTransferTool(jtext, jb_ChooseFile, jb_XmlToExcel, jb_ExcelToXml, oldfilename,
								isExcelToXml).start();
					}
				} else {
					// 获取电脑系统
					System.out.print("" + System.getProperties().getProperty("os.name"));
					JOptionPane.showMessageDialog(jtext, "请重新选择xml或excel文件！");
					jtext.setText("");
					// jb2.setEnabled(false);
					// jb3.setEnabled(false);
				}
			}
		}

		public void dropActionChanged(DropTargetDragEvent dtde) {
			// TODO Auto-generated method stub
			System.out.println("dropActionChanged");
		}

	}

}
