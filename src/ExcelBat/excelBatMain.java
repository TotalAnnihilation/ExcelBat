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
	Font font = new  Font("����",Font.PLAIN, 14);
	JLabel programDescription=new JLabel("    ��������Խ�Ŀ���ļ���ָ������������Դ�ļ���ָ���������ݽ��бȶԣ�");
	JLabel programDescription2=new JLabel("�ȶ���ͬ��Դ�ļ������������Ϣ��ӵ�Ŀ���ļ��У��������ļ�����֧��xls��");
	JLabel jlSourceFile=new JLabel("Դ�ļ�  ��");
	JLabel jlDestinationFile=new JLabel("Ŀ���ļ���");
	JTextField jtSourceFile=new JTextField("��ѡ��ExcelԴ�ļ�");
	JTextField jtDestinationFile=new JTextField("��ѡ��ExcelĿ���ļ�");
	JButton jbSFChose=new JButton("ѡ��");
	JButton jbDFChose=new JButton("ѡ��");
	JLabel jlSFColumn=new JLabel("Դ�ļ��Ա�����  ��");
	JLabel jlDFColumn=new JLabel("Ŀ���ļ��Ա�������");
	JTextField jtSFColumn=new JTextField("����������");
	JTextField jtDFColumn=new JTextField("����������");
//	JFrame jfLoading=new JFrame();
//	JPanel jpLoading =null;
	String strSFColumn=null;
	String strDFColumn=null;
	int iSFColumn =0;
	int iDFColumn =0;
	JLabel jlSFInfoColumn=new JLabel("Դ�ļ�������Ϣ������");
	JTextField jtSFInfoColumn=new JTextField("���������֣��������ÿո����");
	String strSFInfoColumn=null;
	int[] iSFInfoColumn=null;
	JButton jbBAT=new JButton("��ʼ����");
	JLabel jlState=new JLabel("^_^");
	int iState = 0;
	File fileSF=null;
	File fileDF=null;
			
	public excelBatMain()
	{
		try {
			UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
			//ʹ��Windows���Ľ��棬��Ҫ����ļ�ѡ������
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
		this.setTitle("Excel�ļ�������");
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
			//�����ļ�ѡ����ѡ��Դ�ļ���
			JFileChooser jfcSF=new JFileChooser();  
			jfcSF.setDialogTitle("��ѡ��ExcelԴ�ļ�");
			FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel�ļ�(*.xls)", "xls");
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
			//�����ļ�ѡ����ѡ��Ŀ���ļ���
			JFileChooser jfcDF=new JFileChooser();  
			jfcDF.setDialogTitle("��ѡ��ExcelĿ���ļ�");
			FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel�ļ�(*.xls)", "xls");
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
		//���ö��߳̽�����������̣���������ȷ�����̼߳�ʱˢ�½��档
		}
	}
    private  void  excelBat(){
    	//�÷������û�������Ϣ������֤������iState�������׼��״̬��
		iState=0;
		strSFColumn=jtSFColumn.getText();
		try {
		    iSFColumn = Integer.parseInt(strSFColumn)-1;
		} catch (NumberFormatException e1) {
			iState=1;
		    jlState.setText("Դ�ļ��Ա������������");
		}
		strDFColumn=jtDFColumn.getText();
		try {
		    iDFColumn = Integer.parseInt(strDFColumn)-1;
		} catch (NumberFormatException e1) {
			iState=1;
		    jlState.setText("Ŀ���ļ��Ա������������");
		}
		strSFInfoColumn=jtSFInfoColumn.getText();
		//����������������Ϣ�������д������ֻ��Ҫһ�У��򲻰����ո񣬶���������ո���Ҫ�и��ַ�����
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
				    jlState.setText("Դ�ļ�������Ϣ�����������");
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
			    jlState.setText("Դ�ļ�������Ϣ�����������");
			}
		}
		if(fileSF==null){
			iState=1;
			 jlState.setText("��ѡ����ȷ��Դ�ļ���");
		}
		if(fileDF==null){
			iState=1;
			 jlState.setText("��ѡ����ȷ��Ŀ���ļ���");
		}
		if(fileDF.equals(fileSF))
		{
			iState=1;
			 jlState.setText("Դ�ļ���Ŀ���ļ�������ͬ��");
		}
		if(iState==0)
		{
			try {
				 jlState.setText("��ʼ����");
				 //���ݲ�������Excel����
				excelChange(fileSF,fileDF,iSFColumn,iDFColumn,iSFInfoColumn);
				 jlState.setText("����ɹ���");
			} catch (WriteException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
				jlState.setText("����ʧ�ܣ�����δ֪����");
			} catch (BiffException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
				jlState.setText("����ʧ�ܣ�����δ֪����");
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
				jlState.setText("����ʧ�ܣ�����δ֪����");
			}
		}
		
	}
	
	private void excelChange(File fileS,File fileD,int intSFColumn,int intDFColumn,int[] intSFInfoColumn) throws IOException, WriteException, BiffException
	   {
		//�÷�����ɶ�Excel�Ĵ���
	       InputStream streamSF = new FileInputStream(fileS);
	       Workbook wbSF = Workbook.getWorkbook(streamSF);
	       Sheet sheetSF = wbSF.getSheet(0);  
	       InputStream streamDF = new FileInputStream(fileD);
	       Workbook wbDF = Workbook.getWorkbook(streamDF);
	       Sheet sheetDF = wbDF.getSheet(0);  
	       //���ｫ���ļ�����ΪĿ���ļ���_2��
	       File xlsFile = new File(fileD.getAbsolutePath().substring(0,fileD.getAbsolutePath().lastIndexOf('.'))+"_2.xls");
	       // ����һ���ɱ༭�Ĺ�����
	       WritableWorkbook wbNewFile = Workbook.createWorkbook(xlsFile);
	       // ����һ���ɱ༭�Ĺ�����
	       WritableSheet sheetNF = wbNewFile.createSheet("sheet1", 0);
	       //����ļ��Ĳ�����
	       int rowSF = sheetSF.getRows();
	       int rowDF = sheetDF.getRows();
	       int colDF = sheetDF.getColumns();
	       jlState.setText("���ڶ�Ŀ���ļ����д���");
	       //���Ƚ�Ŀ���ļ����е�ȫ�����ݸ��Ƶ����ļ��С�
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
	       //ȷ�����ļ�����е�������������������Դ�ļ����Ƶ����ļ���
	       int[] colNDF=new int[intSFInfoColumn.length];
	       for(int i = 0;i<intSFInfoColumn.length;i++)
	       {
	    	   colNDF[i]=colDF+i;
	    	   Cell TempC = sheetSF.getCell(intSFInfoColumn[i],0);
              	String TempS = TempC.getContents();
              	Label label = new Label(colNDF[i], 0, TempS); 
              	sheetNF.addCell(label);
	       }
	       //��ʼ���бȶԡ�
	       for(int i = 1;i<rowDF;i++)
	       {
	    	   int  ii=i+1;
	    	   jlState.setText("���ڴ����"+ii+"�����ݡ�");
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
	            	   //����ȶԳɹ������н�������Ϣ��Դ�ļ����Ƶ����ļ��С�
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
	       //�������ļ���
	       wbNewFile.write();
	       wbNewFile.close();
	   }
}
