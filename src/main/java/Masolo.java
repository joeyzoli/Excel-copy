import java.awt.BorderLayout;
import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FilenameFilter;

import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;

import com.spire.xls.CellRange;
import com.spire.xls.FileFormat;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;

import javax.swing.JButton;

import sun.misc.Unsafe;

public class Masolo extends JFrame 
{

	private JPanel contentPane;
	private int a = 6;
	private int b = 1;
	private int c = 35;
	private int d = 20;
	private int uja = 6;
	private int ujb = 1;
	private int ujc = 35;
	private int ujd = 20;
	private File mappa;
	private File[] fajlok;
	private Workbook workbook;
	private int szamlalo;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) 
	{
		EventQueue.invokeLater(new Runnable() 
		{
			public void run() 
			{
				try 
				{
					Masolo frame = new Masolo();
					frame.setVisible(true);
				} 
				catch (Exception e) 
				{
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public Masolo() 
	{
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 450, 300);
		setTitle("Telecom Desig másoló");
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		JButton masolgomb = new JButton("Másolás");
		masolgomb.setBounds(164, 89, 89, 23);
		masolgomb.addActionListener(new Masol());
		contentPane.add(masolgomb);
	}
	
	class Masol implements ActionListener
	{
		public void actionPerformed(ActionEvent e) 
		{
			try
			{
				mappa = new File("z:\\Projektek\\Telecom_Design\\5. Végellenőrzés, tesztek\\");
				
				FilenameFilter filter = new FilenameFilter() 											//fájlnév filter metódus
						{
			                public boolean accept(File f, String name) 
			                {
			                    																				// csak az xlsx fájlokat listázza ki 
			                	return name.endsWith(".ods");	
			                }
			            };
			            
				fajlok = mappa.listFiles(filter);
				
				CellRange[] range2 = new CellRange[fajlok.length];
				CellRange[] range4 = new CellRange[fajlok.length];
	            Workbook workbookAll=new Workbook();
				
				for(szamlalo = 0; szamlalo < fajlok.length; szamlalo++)
				{
					if(fajlok[szamlalo].getName().contains("~$"))
					{
						break;
					}
					//Create a Worbook instance
	                workbook = new Workbook();
	                //Load an Excel file
	                workbook.loadFromFile("z:\\Projektek\\Telecom_Design\\5. Végellenőrzés, tesztek\\" + fajlok[szamlalo].getName());
	                //Get the first worksheet
	                Worksheet sheet1 = workbook.getWorksheets().get(0);
	                CellRange range1 = sheet1.getRange().get("A6:T35");		// A6:T35
	                CellRange range3 = sheet1.getRange().get("A2:I2");
	                //range2[szamlalo*2]=range3;
	                range2[szamlalo] = range1;
	                range4[szamlalo] = range3;
	                Worksheet sheetAll = workbookAll.getWorksheets().get(0);
	                
	                int m = range2[szamlalo].getRowCount()*szamlalo;
	                int n = range4[szamlalo].getRowCount()*szamlalo;
	                
	                sheet1.copy(range4[szamlalo], sheetAll.getRange().get(1+m+n,1));
	                sheet1.copy(range2[szamlalo], sheetAll.getRange().get(2+m+n,1));
	                workbookAll.saveToFile("z:\\Babud_Imre\\TelecomOQC\\Telecom_Design_OQC.xlsx");
					System.out.println("Kész");
				}
				JOptionPane.showMessageDialog(null, "Másolás kész!", "Tájékoztató üzenet", 1);
			
			}
			catch(Exception e1)
			{
				System.out.println(fajlok[szamlalo].getName());
				e1.printStackTrace();
				JOptionPane.showMessageDialog(null, "Olvasási hiba történt, az egyik fájl meg van nyitva!", "Hibaüzenet", 2);
			}
			//workbook.saveToFile("c:\\Users\\kovacs.zoltan\\Desktop\\teszt mappa\\masol_teszt.xlsx", FileFormat.Version2013);
			
		}
	}
}
