package ExcelBat;

import java.awt.Font;
import java.awt.Frame;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;


import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.filechooser.FileNameExtensionFilter;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;


public class excelBatMain extends Frame implements ActionListener{
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	Font font = new  Font("宋体",Font.PLAIN, 14);
	JLabel programDescription=new JLabel("    本程序可以将目标文件中指定的列数据与源文件中指定的列数据进行比对，");
	JLabel programDescription2=new JLabel("比对相同则将源文件中所需的列信息添加到目标文件中，生成新文件。仅支持xls！");
	JLabel jlSourceFile=new JLabel("源文件  ：");
	JLabel jlDestinationFile=new JLabel("目标文件：");
	JTextField jtSourceFile=new JTextField("请选择Excel源文件");
	JTextField jtDestinationFile=new JTextField("请选择Excel目标文件");
	JButton jbSFChose=new JButton("选择");
	JButton jbDFChose=new JButton("选择");
	JLabel jlSFColumn=new JLabel("源文件对比列数  ：");
	JLabel jlDFColumn=new JLabel("目标文件对比列数：");
	JTextField jtSFColumn=new JTextField("请输入数字");
	JTextField jtDFColumn=new JTextField("请输入数字");
//	JFrame jfLoading=new JFrame();
//	JPanel jpLoading =null;
	String strSFColumn=null;
	String strDFColumn=null;
	int iSFColumn =0;
	int iDFColumn =0;
	JLabel jlSFInfoColumn=new JLabel("源文件所需信息列数：");
	JTextField jtSFInfoColumn=new JTextField("请输入数字，多列请用空格隔开");
	String strSFInfoColumn=null;
	int[] iSFInfoColumn=null;
	JButton jbBAT=new JButton("开始处理");
	JLabel jlState=new JLabel("^_^");
	int iState = 0;
	File fileSF=null;
	File fileDF=null;
			
	public excelBatMain()
	{
		try {
			UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
			//使用Windows风格的界面，主要针对文件选择器。
		} catch (ClassNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (InstantiationException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (IllegalAccessException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (UnsupportedLookAndFeelException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		this.setTitle("Excel文件批处理");
		this.setLayout(null);
		programDescription.setFont(font);
		programDescription.setBounds(50, 50, 550, 20);
		this.add(programDescription);
		programDescription2.setFont(font);
		programDescription2.setBounds(50, 75, 550, 20);
		this.add(programDescription2);
		jlSourceFile.setFont(font);
		jlSourceFile.setBounds(50, 120, 100, 20);
		this.add(jlSourceFile);
		jtSourceFile.setFont(font);
		jtSourceFile.setBounds(150, 120, 300, 20);
		this.add(jtSourceFile);
		jlDestinationFile.setFont(font);
		jlDestinationFile.setBounds(50, 170, 100, 20);
		this.add(jlDestinationFile);
		jtDestinationFile.setFont(font);
		jtDestinationFile.setBounds(150, 170, 300, 20);
		this.add(jtDestinationFile);
		jbSFChose.setFont(font);
		jbSFChose.setBounds(500, 120, 80, 20);
		this.add(jbSFChose);
		jbSFChose.addActionListener(this);
		jbDFChose.setFont(font);
		jbDFChose.setBounds(500, 170, 80, 20);
		this.add(jbDFChose);
		jbDFChose.addActionListener(this);
		
		jlSFColumn.setFont(font);
		jlSFColumn.setBounds(50, 220, 150, 20);
		this.add(jlSFColumn);
		jtSFColumn.setFont(font);
		jtSFColumn.setBounds(230, 220, 100, 20);
		this.add(jtSFColumn);
		jlDFColumn.setFont(font);
		jlDFColumn.setBounds(50, 270, 150, 20);
		this.add(jlDFColumn);
		jtDFColumn.setFont(font);
		jtDFColumn.setBounds(230, 270, 100, 20);
		this.add(jtDFColumn);
		jlSFInfoColumn.setFont(font);
		jlSFInfoColumn.setBounds(50, 320, 150, 20);
		this.add(jlSFInfoColumn);
		jtSFInfoColumn.setFont(font);
		jtSFInfoColumn.setBounds(230, 320, 300, 20);
		this.add(jtSFInfoColumn);
		jbBAT.setFont(font);
		jbBAT.setBounds(160, 370, 100, 20);
		this.add(jbBAT);
		jlState.setBounds(300, 370, 200, 20);
		jlState.setFont(font);
		this.add(jlState);
		jbBAT.addActionListener(this);
		this.setBounds(300, 100, 650, 450);
		this.setVisible(true);
		this.addWindowListener(new WindowAdapter()
		{
		   public void windowClosing(WindowEvent e)
		   {
			   super.windowClosing(e);
			   System.exit(0);
		   }
		});
	}
	public static void main(String args[])
	{
		new excelBatMain();
	}
	public void actionPerformed(ActionEvent e)
	{
		if(e.getSource()==jbSFChose)
		{
			//利用文件选择器选择源文件。
			JFileChooser jfcSF=new JFileChooser();  
			jfcSF.setDialogTitle("请选择Excel源文件");
			FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel文件(*.xls)", "xls");
		    jfcSF.setFileFilter(filter);
		    int returnVal =  jfcSF.showOpenDialog(null);
		    if (returnVal == JFileChooser.APPROVE_OPTION) 
		    {
		    	fileSF=jfcSF.getSelectedFile(); 
		    	jtSourceFile.setText(fileSF.getAbsolutePath());
		    }
		}
		if(e.getSource()==jbDFChose)
		{
			//利用文件选择器选择目标文件。
			JFileChooser jfcDF=new JFileChooser();  
			jfcDF.setDialogTitle("请选择Excel目标文件");
			FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel文件(*.xls)", "xls");
		    jfcDF.setFileFilter(filter);
		    int returnVal =  jfcDF.showOpenDialog(null);
		    if (returnVal == JFileChooser.APPROVE_OPTION) 
		    {
		    	fileDF=jfcDF.getSelectedFile(); 
		    	jtDestinationFile.setText(fileDF.getAbsolutePath());
		    }
		}
		if(e.getSource()==jbBAT)
		{
		new Thread(new Runnable() {
	            @Override
	            public void run() {
	            	excelBat();
	            }}).start();
		//利用多线程进行批处理过程，这样可以确保主线程及时刷新界面。
		}
	}
    private  void  excelBat(){
    	//该方法对用户所填信息进行验证。利用iState变量标记准备状态。
		iState=0;
		strSFColumn=jtSFColumn.getText();
		try {
		    iSFColumn = Integer.parseInt(strSFColumn)-1;
		} catch (NumberFormatException e1) {
			iState=1;
		    jlState.setText("源文件对比列数输入错误！");
		}
		strDFColumn=jtDFColumn.getText();
		try {
		    iDFColumn = Integer.parseInt(strDFColumn)-1;
		} catch (NumberFormatException e1) {
			iState=1;
		    jlState.setText("目标文件对比列数输入错误！");
		}
		strSFInfoColumn=jtSFInfoColumn.getText();
		//下面对输入的所需信息列数进行处理，如果只需要一列，则不包含空格，多列则包含空格，需要切割字符串。
		if(strSFInfoColumn.indexOf(" ")!=-1)
		{
			String str[] = strSFInfoColumn.split(" ");  
			int[] iTemp1=new int[str.length];
			for(int i=0;i<str.length;i++)
			{
				try {
					iTemp1[i]=Integer.parseInt(str[i])-1;  
				} catch (NumberFormatException e2) {
					iState=1;
				    jlState.setText("源文件所需信息列数输入错误！");
				    break;
				}
			}
			iSFInfoColumn=iTemp1;
		} 
		else
		{
			try {
				int[] iTemp2=new int [1];
				iTemp2[0] = Integer.parseInt(strSFInfoColumn)-1;
				iSFInfoColumn=iTemp2;
			} catch (NumberFormatException e1) {
				iState=1;
			    jlState.setText("源文件所需信息列数输入错误！");
			}
		}
		if(fileSF==null){
			iState=1;
			 jlState.setText("请选择正确的源文件！");
		}
		if(fileDF==null){
			iState=1;
			 jlState.setText("请选择正确的目标文件！");
		}
		if(fileDF.equals(fileSF))
		{
			iState=1;
			 jlState.setText("源文件与目标文件不能相同！");
		}
		if(iState==0)
		{
			try {
				 jlState.setText("开始处理！");
				 //传递参数进行Excel处理。
				excelChange(fileSF,fileDF,iSFColumn,iDFColumn,iSFInfoColumn);
				 jlState.setText("处理成功！");
			} catch (WriteException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
				jlState.setText("处理失败，发生未知错误！");
			} catch (BiffException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
				jlState.setText("处理失败，发生未知错误！");
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
				jlState.setText("处理失败，发生未知错误！");
			}
		}
		
	}
	
	private void excelChange(File fileS,File fileD,int intSFColumn,int intDFColumn,int[] intSFInfoColumn) throws IOException, WriteException, BiffException
	   {
		//该方法完成对Excel的处理。
	       InputStream streamSF = new FileInputStream(fileS);
	       Workbook wbSF = Workbook.getWorkbook(streamSF);
	       Sheet sheetSF = wbSF.getSheet(0);  
	       InputStream streamDF = new FileInputStream(fileD);
	       Workbook wbDF = Workbook.getWorkbook(streamDF);
	       Sheet sheetDF = wbDF.getSheet(0);  
	       //这里将新文件保存为目标文件名_2。
	       File xlsFile = new File(fileD.getAbsolutePath().substring(0,fileD.getAbsolutePath().lastIndexOf('.'))+"_2.xls");
	       // 创建一个可编辑的工作簿
	       WritableWorkbook wbNewFile = Workbook.createWorkbook(xlsFile);
	       // 创建一个可编辑的工作表
	       WritableSheet sheetNF = wbNewFile.createSheet("sheet1", 0);
	       //获得文件的参数。
	       int rowSF = sheetSF.getRows();
	       int rowDF = sheetDF.getRows();
	       int colDF = sheetDF.getColumns();
	       jlState.setText("正在对目标文件进行处理！");
	       //首先将目标文件已有的全部内容复制到新文件中。
	       for(int i = 0;i<rowDF;i++)
	       {
	    	   for (int j=0;j<colDF;j++)
	    	   {
	    		   Cell TempC = sheetDF.getCell(j,i);
        		   String TempS = TempC.getContents();
        		   Label label = new Label(j,i,TempS); 
        		   sheetNF.addCell(label);
	    	   }
	       }
	       //确定新文件添加列的列数，并将标题行由源文件复制到新文件。
	       int[] colNDF=new int[intSFInfoColumn.length];
	       for(int i = 0;i<intSFInfoColumn.length;i++)
	       {
	    	   colNDF[i]=colDF+i;
	    	   Cell TempC = sheetSF.getCell(intSFInfoColumn[i],0);
              	String TempS = TempC.getContents();
              	Label label = new Label(colNDF[i], 0, TempS); 
              	sheetNF.addCell(label);
	       }
	       //开始进行比对。
	       for(int i = 1;i<rowDF;i++)
	       {
	    	   int  ii=i+1;
	    	   jlState.setText("正在处理第"+ii+"行数据。");
	    	   Cell cDF = sheetDF.getCell(intDFColumn, i);
	           String strDF = cDF.getContents();
//	           System.out.println("I="+i);
	           for (int j=1;j<rowSF;j++)
	           {
	        	   Cell cSF = sheetSF.getCell(intSFColumn, j);
	               String strSF = cSF.getContents();
//	           System.out.println("J="+j);
//	           System.out.println(strDF);
//	           System.out.println(strSF);
	               if(strDF.equals(strSF))
	               {
	            	   //如果比对成功，逐列将所需信息从源文件复制到新文件中。
	            	   for(int k = 0;k<intSFInfoColumn.length;k++)
	            	   {
	            		   Cell TempC = sheetSF.getCell(intSFInfoColumn[k], j);
	            		   String TempS = TempC.getContents();
	            		   Label label = new Label(colNDF[k], i, TempS); 
	            		   sheetNF.addCell(label);
//	            		   System.out.println("K="+k);
	            	   }
	            	   break;
	               }
	           }
	       }
	       //保存新文件。
	       wbNewFile.write();
	       wbNewFile.close();
	   }
}
