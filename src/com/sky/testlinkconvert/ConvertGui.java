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
 * testlinkconvert的图形界面
 * @author Rachel.Luo
 * */
public class ConvertGui implements ActionListener{
	
	JTextField ja1 = new DropDragSupportTextArea(20);
	FileDialog fd = null;
	JFrame jf = null;
	JButton jb1 =null;
	JButton jb2 = null;
	JButton jb3 = null;
	String oldfilename;
	boolean isExcelToXml=false;
	
    public ConvertGui(){
    	jf = new JFrame("Testlink转换器");
    	fd = new FileDialog(jf);
    	JPanel j1 = new JPanel();
    	JPanel j2 = new JPanel();
    	JLabel jl1 = new JLabel("源文件:");
    	
    	jb1 = new JButton("选择");
    	jb2 = new JButton("xml转成excel");
    	jb3 = new JButton("excel转成xml");
    	jb1.addActionListener(this);
    	jb2.addActionListener(this);
    	jb3.addActionListener(this);
    	j1.add(jl1);
    	ja1.setEditable(false);
    	ja1.setBackground(Color.white);
    	ja1.addActionListener(this);
    	j1.add(ja1);
    	j1.add(jb1);
    	jb2.setEnabled(false);
    	jb3.setEnabled(false);
    	j2.add(jb2);
    	j2.add(jb3);
    	jf.add(j1,"North");
    	jf.add(j2);
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
		if(comm.equals("选择")){
			fd.setVisible(true);
			if(fd.getFile()!=null){
				if(fd.getFile().endsWith(".xml")||fd.getFile().endsWith(".xls")
						||fd.getFile().endsWith(".xlsx")){
					ja1.setText(fd.getDirectory()+fd.getFile());
					oldfilename=fd.getDirectory()+fd.getFile();
					
					if(fd.getFile().endsWith(".xml")){
						jb2.setEnabled(true);
						jb3.setEnabled(false);
						isExcelToXml=false;
						jb2.setEnabled(false);
						jb1.setEnabled(false);
						new FileTransferTool(ja1,jb1,jb2,jb3,oldfilename,isExcelToXml).start();
					}else{
						jb3.setEnabled(true);
						jb2.setEnabled(false);
						isExcelToXml=true;
						jb3.setEnabled(false);
						jb1.setEnabled(false);
						new FileTransferTool(ja1,jb1,jb2,jb3,oldfilename,isExcelToXml).start();
					}
				}else{
					System.out.print(""+System.getProperties().getProperty("os.name"));
					JOptionPane.showMessageDialog(ja1,"请重新选择xml或excel文件！");
					ja1.setText("");
					jb2.setEnabled(false);
					jb3.setEnabled(false);
				}
			}
		}else{
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
			dropTarget = new DropTarget(this, DnDConstants.ACTION_COPY_OR_MOVE,
					this, true);
		}

		/**
		 * 拖入文件或字符串,这里只说明能拖拽，并未打开文件并显示到文本区域中
		 */
		public void dragEnter(DropTargetDragEvent dtde) {
			DataFlavor[] dataFlavors = dtde.getCurrentDataFlavors();
			if (dataFlavors[0].match(DataFlavor.javaFileListFlavor)) {
				try {
					Transferable tr = dtde.getTransferable();
					Object obj = tr.getTransferData(DataFlavor.javaFileListFlavor);
					List<File> files = (List<File>) obj;
				  
					if (files != null &&files.size() > 0){
						String absolutePath = files.get(0).getAbsolutePath();
						ja1.setText(absolutePath);					
					}
					/*
					for (int i = 0; i < files.size(); i++) {
						append(files.get(i).getAbsolutePath() + "/r/n");
					}*/
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
			if(ja1!=null){
				System.out.println("开始");
				if(ja1.getText().endsWith(".xml")||ja1.getText().endsWith(".xls")
						||ja1.getText().endsWith(".xlsx")){
					//ja1.setText(fd.getDirectory()+fd.getFile());
					//oldfilename=fd.getDirectory()+fd.getFile();
					oldfilename = ja1.getText();
					if(ja1.getText().endsWith(".xml")){
						jb2.setEnabled(true);
						jb3.setEnabled(false);
						isExcelToXml=false;
						jb2.setEnabled(false);
						jb1.setEnabled(false);
						new FileTransferTool(ja1,jb1,jb2,jb3,oldfilename,isExcelToXml).start();
					}else{
						jb3.setEnabled(true);
						jb2.setEnabled(false);
						isExcelToXml=true;
						jb3.setEnabled(false);
						jb1.setEnabled(false);
						new FileTransferTool(ja1,jb1,jb2,jb3,oldfilename,isExcelToXml).start();
					}
				}else{
					//获取电脑系统
					System.out.print(""+System.getProperties().getProperty("os.name"));
					JOptionPane.showMessageDialog(ja1,"请重新选择xml或excel文件！");
					ja1.setText("");
					//jb2.setEnabled(false);
					//jb3.setEnabled(false);
				}
			}
		}

		public void dropActionChanged(DropTargetDragEvent dtde) {
			// TODO Auto-generated method stub
			System.out.println("dropActionChanged");
		}

		
	}
	
	
}


