package com.sky.testlinkconvert;

import javax.swing.*;

/**
 * 实际转换文件的线程类
 * 
 * */
public class FileTransferTool extends Thread{
	private JTextField jText = null;
	private JButton jb_ChooseFile = null;
	private JButton jb_XmlToExcel = null;
	private JButton jb_ExcelToXml = null;
	private String oldfilename;
	private boolean isExcelToXml;
	
    public FileTransferTool(JTextField ja1,JButton jb1,JButton jb2,JButton jb3,String oldfilename,boolean isExcelToXml){
    	this.jText=ja1;
    	this.jb_ChooseFile=jb1;
    	this.jb_XmlToExcel=jb2;
    	this.jb_ExcelToXml=jb3;
    	this.oldfilename=oldfilename;
    	this.isExcelToXml=isExcelToXml;
    }
    
	@Override
	public void run() {
		if(isExcelToXml!=true){
			//JOptionPane.showMessageDialog(jb2,"文件转换中，请点击确定，等待完成提示...");
			System.out.println("xml to excel convert start!");
			System.out.println("oldfilename:"+oldfilename);
			XmlToExcel.transferXMLToExcel(oldfilename);
			JOptionPane.showMessageDialog(jb_XmlToExcel,"文件转换完成，请到源文件目录查看！");
			System.out.println("xml to excel convert end!");
			jText.setText("");
			jb_ChooseFile.setEnabled(true);
			
		}else{
			//JOptionPane.showMessageDialog(jb3,"文件转换中，请点击确定，等待完成提示...");
			System.out.println("excel to xml convert start!");
			System.out.println("oldfilename:"+oldfilename);
			ExcelToXml.transferExcelToXml(oldfilename);
			JOptionPane.showMessageDialog(jb_ExcelToXml,"文件转换完成，请到源文件目录查看！");
			System.out.println("excel to xml convert end!");
			jText.setText("");
			jb_ChooseFile.setEnabled(true);
			
		}
	}
}
