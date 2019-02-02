import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.border.EmptyBorder;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

public class FormController implements ActionListener {
	
	private DataEntryForm form;
	private JFrame cashFrame;
	private JTextField jt1c;
	private JTextField jt5c;
	private JTextField jt10c;
	private JTextField jt25c;
	private JTextField jt50c;
	private JTextField jt1;
	private JTextField jt2;
	private JTextField jt5;
	private JTextField jt10;
	private JTextField jt20;
	private JTextField jt50;
	private JTextField jt100;
	
	private static ArrayList<Donation> offering = new ArrayList<Donation>();

	public FormController(DataEntryForm f) {
		super();
		form = f;
		setupCashFrame();
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		
		System.out.println("event occured - " + e.getActionCommand());
		
		if(e.getActionCommand().compareTo("enter-data")==0) {
			try{
				enterData();
			}catch(Exception e1){
				System.out.println(e1);
				JOptionPane.showMessageDialog(null, "There was a problem processing your entires.  \nLikely you did not enter a number in the amount field.  Please, look them over.");
			}
		}
			
		if(e.getActionCommand().compareTo("show-data")==0)
			showAllEntries();
		
		if(e.getActionCommand().compareTo("envelope-event")==0)
			envelopeEvent();
		
		if(e.getActionCommand().compareTo("lastname-event")==0)
			lastNameEvent();
		
		if(e.getActionCommand().compareTo("exit-program")==0)
			exitProgram();
		
		if(e.getActionCommand().compareTo("enter-cash")==0) {
			enterCash();
		}
	}
	
	private void enterCash() {
		 
		 cashFrame.setVisible(true);
		 
	}
	
	private void setupCashFrame() {
		cashFrame = new JFrame("Start up frame");		 

		 cashFrame.setDefaultCloseOperation(JFrame.HIDE_ON_CLOSE);
		 cashFrame.setBounds(200, 120, 300, 460);
		 
		 JPanel pane = new JPanel();
		 pane.setBorder(new EmptyBorder(5, 5, 5, 5));
		 pane.setBackground(new Color(62, 100, 124));
		 pane.setLayout(null);
		 
		 JButton enterCashButton = new JButton("Enter");
		 enterCashButton.setBounds(70, 400, 80, 20);
		 pane.add(enterCashButton);
		 
		 enterCashButton.addActionListener(new ActionListener() {
		        @Override
		        public void actionPerformed(ActionEvent e) {
		        		System.out.println("in cash button action event");
		        		System.out.println(jt1c.getText());
		        		cashFrame.setVisible(false);
		        		exportToExcel();
		        }
		    });
		 
		 JButton cancelButton = new JButton("Cancel");
		 cancelButton.setBounds(170, 400, 80, 20);
		 pane.add(cancelButton);
		 
		 ///////////////////////////////////////////
		 
		 JLabel label100 = new JLabel("$100.00");
		 label100.setHorizontalAlignment(SwingConstants.RIGHT);
		 label100.setBounds(50, 30, 80, 20);
		 label100.setForeground(Color.white);
		 pane.add(label100);
		 
		 jt100 = new JTextField();
		 jt100.setBounds(150, 30, 80, 20);
		 jt100.setText("0");
		 pane.add(jt100);
		 jt100.setColumns(10);
		 
		 //////////////////////////////////////////////
		 
		 JLabel label50 = new JLabel("$50.00");
		 label50.setHorizontalAlignment(SwingConstants.RIGHT);
		 label50.setBounds(50, 60, 80, 20);
		 label50.setForeground(Color.white);
		 pane.add(label50);
		 
		 jt50 = new JTextField();
		 jt50.setBounds(150, 60, 80, 20);
		 jt50.setText("0");
		 pane.add(jt50);
		 jt50.setColumns(10);
		 
		 //////////////////////////////////////////////
		 
		 JLabel label20 = new JLabel("$20.00");
		 label20.setHorizontalAlignment(SwingConstants.RIGHT);
		 label20.setBounds(50, 90, 80, 20);
		 label20.setForeground(Color.white);
		 pane.add(label20);
		 
		 jt20 = new JTextField();
		 jt20.setBounds(150, 90, 80, 20);
		 jt20.setText("0");
		 pane.add(jt20);
		 jt20.setColumns(10);
		 
		 //////////////////////////////////////////////
		 
		 JLabel label10 = new JLabel("$10.00");
		 label10.setHorizontalAlignment(SwingConstants.RIGHT);
		 label10.setBounds(50, 120, 80, 20);
		 label10.setForeground(Color.white);
		 pane.add(label10);
		 
		 jt10 = new JTextField();
		 jt10.setBounds(150, 120, 80, 20);
		 jt10.setText("0");
		 pane.add(jt10);
		 jt10.setColumns(10);
		 
		 //////////////////////////////////////////////
		 
		 JLabel label5 = new JLabel("$5.00");
		 label5.setHorizontalAlignment(SwingConstants.RIGHT);
		 label5.setBounds(50, 150, 80, 20);
		 label5.setForeground(Color.white);
		 pane.add(label5);
		 
		 jt5 = new JTextField();
		 jt5.setBounds(150, 150, 80, 20);
		 jt5.setText("0");
		 pane.add(jt5);
		 jt5.setColumns(10);
		 
		 //////////////////////////////////////////////
		 
		 JLabel label2 = new JLabel("$2.00");
		 label2.setHorizontalAlignment(SwingConstants.RIGHT);
		 label2.setBounds(50, 180, 80, 20);
		 label2.setForeground(Color.white);
		 pane.add(label2);
		 
		 jt2 = new JTextField();
		 jt2.setBounds(150, 180, 80, 20);
		 jt2.setText("0");
		 pane.add(jt2);
		 jt2.setColumns(10);
		 
		 //////////////////////////////////////////////
		 
		 JLabel label1 = new JLabel("$1.00");
		 label1.setHorizontalAlignment(SwingConstants.RIGHT);
		 label1.setBounds(50, 210, 80, 20);
		 label1.setForeground(Color.white);
		 pane.add(label1);
		 
		 jt1 = new JTextField();
		 jt1.setBounds(150, 210, 80, 20);
		 jt1.setText("0");
		 pane.add(jt1);
		 jt1.setColumns(10);
		 
		 //////////////////////////////////////////////
		 
		 JLabel label50c = new JLabel("$0.50");
		 label50c.setHorizontalAlignment(SwingConstants.RIGHT);
		 label50c.setBounds(50, 240, 80, 20);
		 label50c.setForeground(Color.white);
		 pane.add(label50c);
		 
		 jt50c = new JTextField();
		 jt50c.setBounds(150, 240, 80, 20);
		 jt50c.setText("0");
		 pane.add(jt50c);
		 jt50c.setColumns(10);
		 
		 //////////////////////////////////////////////
		 
		 JLabel label25c = new JLabel("$0.25");
		 label25c.setHorizontalAlignment(SwingConstants.RIGHT);
		 label25c.setBounds(50, 270, 80, 20);
		 label25c.setForeground(Color.white);
		 pane.add(label25c);
		 
		 jt25c = new JTextField();
		 jt25c.setBounds(150, 270, 80, 20);
		 jt25c.setText("0");
		 pane.add(jt25c);
		 jt25c.setColumns(10);
		 
		 //////////////////////////////////////////////
		 
		 JLabel label10c = new JLabel("$0.10");
		 label10c.setHorizontalAlignment(SwingConstants.RIGHT);
		 label10c.setBounds(50, 300, 80, 20);
		 label10c.setForeground(Color.white);
		 pane.add(label10c);
		 
		 jt10c = new JTextField();
		 jt10c.setBounds(150, 300, 80, 20);
		 jt10c.setText("0");
		 pane.add(jt10c);
		 jt10c.setColumns(10);
		 
		 //////////////////////////////////////////////
		 
		 JLabel label5c = new JLabel("$0.05");
		 label5c.setHorizontalAlignment(SwingConstants.RIGHT);
		 label5c.setBounds(50, 330, 80, 20);
		 label5c.setForeground(Color.white);
		 pane.add(label5c);
		 
		 jt5c = new JTextField();
		 jt5c.setBounds(150, 330, 80, 20);
		 jt5c.setText("0");
		 pane.add(jt5c);
		 jt5c.setColumns(10);
		 
		 //////////////////////////////////////////////
		 
		 JLabel label1c = new JLabel("$0.01");
		 label1c.setHorizontalAlignment(SwingConstants.RIGHT);
		 label1c.setBounds(50, 360, 80, 20);
		 label1c.setForeground(Color.white);
		 pane.add(label1c);
		 
		 jt1c = new JTextField();
		 jt1c.setBounds(150, 360, 80, 20);
		 jt1c.setText("0");
		 pane.add(jt1c);
		 jt1c.setColumns(10);
		 
		 ///////////////////////////////////////////////
		
		 cashFrame.add(pane);
	}
	
	public void exitProgram() {
		System.exit(0);
	}
		
	private void lastNameEvent() {
		
		// Force first letter to be capitalized
		String temp = capitalizeString(form.getLastNameField().getText());
		form.getLastNameField().setText(temp);
		
		// the array implementation which leads to a scroll box
		String[] namesList = new String [form.getChurchDB().size()];
		
		int i=0;
		for( Donor d: form.getChurchDB()) {
			String searchName = form.getLastNameField().getText();
			if(d.getLastName().length()>=searchName.length()) {
				if(d.getLastName().substring(0,searchName.length()).compareToIgnoreCase(searchName)==0) {
					namesList[i] = d.getLastName()+", " + d.getFirstName()+"  ";
					i++;
				}
			}
		}
		
		// name not in database
		if (i==0) {
			resetNonLastNameFields();
			JOptionPane.showMessageDialog(null,"no matches or partial matches for " + form.getLastNameField().getText() + " found in church database");
		}  // only one name match
		else if(i==1) {
			boolean match=false;
			
			for(Donor d: form.getChurchDB()) {
				if (d.getLastName().compareToIgnoreCase(form.getLastNameField().getText())==0) {
					form.getLastNameField().setText(d.getLastName());
					form.getFirstNameField().setText(d.getFirstName());
					form.getEnvelopeField().setText(d.getEnvelopeNumber());
					form.getAddressField().setText(d.getAddress());
					form.getCityField().setText(d.getCity());
					form.getStateField().setSelectedItem(d.getState());
					form.getZipField().setText(d.getZip());
					match=true;
				}
			}
			if(!match) {
				String n = (String)JOptionPane.showInputDialog(null, "Select a person ",
		                "names", JOptionPane.QUESTION_MESSAGE, null, namesList, namesList[0]);

					if(n!=null) {
						for(Donor d: form.getChurchDB()) {
							if(d.getLastName().compareToIgnoreCase(n.substring(0, n.indexOf(',')))==0)
								if(d.getFirstName().compareToIgnoreCase(n.substring(n.indexOf(',')+2, n.length()-2))==0){
								form.getLastNameField().setText(d.getLastName());
								form.getFirstNameField().setText(d.getFirstName());
								form.getEnvelopeField().setText(d.getEnvelopeNumber());
								form.getAddressField().setText(d.getAddress());
								form.getCityField().setText(d.getCity());
								form.getStateField().setSelectedItem(d.getState());
								form.getZipField().setText(d.getZip());
								break;
							}
						}
					}
			}
		}
		 else{
		
			String n = (String)JOptionPane.showInputDialog(null, "Select a person ",
                "names", JOptionPane.QUESTION_MESSAGE, null, namesList, namesList[0]);

			if(n!=null) {
				for(Donor d: form.getChurchDB()) {
					if(d.getLastName().compareToIgnoreCase(n.substring(0, n.indexOf(',')))==0)
						if(d.getFirstName().compareToIgnoreCase(n.substring(n.indexOf(',')+2, n.length()-2))==0){
						form.getLastNameField().setText(d.getLastName());
						form.getFirstNameField().setText(d.getFirstName());
						form.getEnvelopeField().setText(d.getEnvelopeNumber());
						form.getAddressField().setText(d.getAddress());
						form.getCityField().setText(d.getCity());
						form.getStateField().setSelectedItem(d.getState());
						form.getZipField().setText(d.getZip());
						break;
					}
				}
			}
		}
	}
			
	private void envelopeEvent() {
		String en = (String)form.getEnvelopeField().getText();
		boolean envelopeExists=false;
		
		for(Donor d1: form.getChurchDB()) {
			if(d1.getEnvelopeNumber()!=null) 
				if (d1.getEnvelopeNumber().compareTo(en)==0)
					envelopeExists=true;
		}
		if(envelopeExists) {
			fillInDataUsingEnvelopeNumber(en);
		}
		else {
			JOptionPane.showMessageDialog(form.getContentPane(), "Envelope number " + form.getEnvelopeField().getText() + " does not exist in database", 
					"Data Entry Problem Message", JOptionPane.ERROR_MESSAGE);
		}	
	}
	
	private void enterData() {

		System.out.println("passed 0");
		
		if(form.getAmountField().getText().compareToIgnoreCase("")==0) {
			JOptionPane.showMessageDialog(null, "No value entered for amount.");
			return;
		}
		
		Donation d = getFormData();
		
		// ask to add the person to the database if they are not in it already
		if(!isNameInDB()) {
			int choice = JOptionPane.showOptionDialog(null, 
				      "The person is not in the database, would you like to add them?", 
				      "Add person to church database?", 
				      JOptionPane.YES_NO_OPTION, 
				      JOptionPane.QUESTION_MESSAGE, 
				      null, null, null);
			// the YES choice
			if(choice==0) {
				form.getChurchDB().add(d.getDonor());
				updateChurchDB();
			}
		}
		
		// a==2 complete name and address match found
		// a==1 name match found, but address is different
		// a==0 name not found in db
		int  a=checkForAddressMatchInDB();
		if(a==1) {
			int choice = JOptionPane.showOptionDialog(null, 
				      form.getLastNameField().getText() + ", " + form.getFirstNameField().getText() + " is in the database, but address entered does not match the database information.  \nWould you like to update the databaase?", 
				      "Update Address church database?", 
				      JOptionPane.YES_NO_OPTION, 
				      JOptionPane.QUESTION_MESSAGE, 
				      null, null, null);
			if(choice==0) {
				Donor d1 = new Donor();
				d1.setLastName(form.getLastNameField().getText());
				d1.setFirstName(form.getFirstNameField().getText());
				d1.setEnvelopeNumber(form.getEnvelopeField().getText());
				d1.setAddress(form.getAddressField().getText());
				d1.setCity(form.getCityField().getText());
				d1.setState(form.getStateField().getSelectedItem().toString());
				d1.setZip(form.getZipField().getText());
				for(int i=0; i<form.getChurchDB().size(); i++) {
					Donor dReplace = form.getChurchDB().get(i);
					if(dReplace.getFirstName().equals(d1.getFirstName()) && dReplace.getLastName().equals(d1.getLastName())) {
						form.getChurchDB().set(i,d1);
						break;
					}
				}			
				updateChurchDB();
			}
		}
				
		if(!isLegalEnvelopeEntry()) {
			JOptionPane.showMessageDialog(null, "Entries designated as envelope must have an envelope #");
			return;
		}
		
		addDonationToOffering(d);
		// check to see if this is a new description field
		if(d.getDescription()!=null) {
			boolean isInList=false;
			for(int i=0; i<form.getDescriptionField().getItemCount(); i++) {
				if(d.getDescription().compareToIgnoreCase(form.getDescriptionField().getItemAt(i))==0)
					isInList=true;
			}
			if(!isInList)
				form.getDescriptionField().addItem(d.getDescription());
		}
		
		System.out.println("passed 6");
		
		form.getLastNameField().setActionCommand("lastname-event-off");
		
		System.out.println("passed 7");
		
		resetForm();
		
		System.out.println("passed 8");
		
		form.getLastNameField().setActionCommand("lastname-event");
		
		System.out.println("passed 9");
		
		exportToExcel();
		
		System.out.println("passed 10");
	}
	
	private boolean isLegalEnvelopeEntry() {
		if(form.getDesignationField().getSelectedItem().toString().compareToIgnoreCase("envelope")==0)
			if(form.getEnvelopeField().getText().compareToIgnoreCase("")==0)
				return false;
			return true;
	}
	
	private boolean isNameInDB() {
		boolean match = false;
		for(Donor d: form.getChurchDB()) {
			if (d.getLastName().compareToIgnoreCase(form.getLastNameField().getText())==0)
				if(d.getFirstName().compareToIgnoreCase(form.getFirstNameField().getText())==0)
					match=true;
		}
		return match;
	}
	
	public void addDonationToOffering(Donation d) {
		offering.add(d);
	}
	
	private void resetForm() {
		//updateLastNameComboBox("");
		form.getLastNameField().setText("");
		form.getFirstNameField().setText("");
		form.getEnvelopeField().setText("");
		form.getAddressField().setText("");
		form.getCityField().setText("");
		form.getZipField().setText("");
		form.getStateField().setSelectedIndex(0);
		form.getCategoryField().setSelectedIndex(0);
		form.getDesignationField().setSelectedIndex(0);
		form.getDescriptionField().setSelectedIndex(0);
		form.getAmountField().setText("");
	}
	
	private String capitalizeString(String s) {
		return s.substring(0, 1).toUpperCase() + s.substring(1);
	}
	
	private void resetNonLastNameFields() {
		form.getFirstNameField().setText("");
		form.getEnvelopeField().setText("");
		form.getAddressField().setText("");
		form.getCityField().setText("");
		form.getZipField().setText("");
		form.getStateField().setSelectedIndex(0);
	}
	
	private void updateChurchDB() {

		String excelFileName = "churchDB.xls"; //name of excel file

		String sheetName = "Sheet1"; //name of sheet

		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet(sheetName) ;

		HSSFRow row = sheet.createRow(0);
		
		HSSFCell cell = row.createCell(0);
		cell.setCellValue("Env #");
		
		cell = row.createCell(1);
		cell.setCellValue("Last Name");
		
		cell = row.createCell(2);
		cell.setCellValue("First Name");
		
		cell = row.createCell(3);
		cell.setCellValue("Address");
		
		cell = row.createCell(4);
		cell.setCellValue("City");
		
		cell = row.createCell(5);
		cell.setCellValue("State");
		
		cell = row.createCell(6);
		cell.setCellValue("Zip");

		Collections.sort(form.getChurchDB(), new CustomComparator());
		
		//iterating r number of rows
		for (int r=0;r < form.getChurchDB().size(); r++ )
		{
			row = sheet.createRow(r+1);
			
			cell = row.createCell(0);
			cell.setCellValue(form.getChurchDB().get(r).getEnvelopeNumber());
			
			cell = row.createCell(1);
			cell.setCellValue(form.getChurchDB().get(r).getLastName());
			
			cell = row.createCell(2);
			cell.setCellValue(form.getChurchDB().get(r).getFirstName());
			
			cell = row.createCell(3);
			cell.setCellValue(form.getChurchDB().get(r).getAddress());
			
			cell = row.createCell(4);
			cell.setCellValue(form.getChurchDB().get(r).getCity());
			
			cell = row.createCell(5);
			cell.setCellValue(form.getChurchDB().get(r).getState());
			
			cell = row.createCell(6);
			cell.setCellValue(form.getChurchDB().get(r).getZip());
	
		}
		try {
			FileOutputStream fileOut = new FileOutputStream(excelFileName);
		
			//write this workbook to an Outputstream.
			wb.write(fileOut);
			wb.close();
			fileOut.flush();
			fileOut.close();
		}
		catch(Exception e) {
			System.out.println(e);
		}
	}


	public String nullToEmptyString(String s) {
		if (s==null)
			return "";
		else 
			return s;
	}
	
	public Integer valueOf(String s) {
		if (s==null  || s.compareTo("")==0)
			return 0;
		else 
			return Integer.parseInt(s);
	}
	
	public Donation getFormData() {

		Donor d = new Donor(
				nullToEmptyString((String)form.getLastNameField().getText()),
				nullToEmptyString(form.getFirstNameField().getText()),
				nullToEmptyString(form.getEnvelopeField().getText()),
				nullToEmptyString(form.getAddressField().getText()),
				nullToEmptyString(form.getCityField().getText()),
				nullToEmptyString((String)form.getStateField().getSelectedItem()),
				nullToEmptyString(form.getZipField().getText())
				);
		
		Donation s = new Donation(
				d,
				nullToEmptyString((String)form.getCategoryField().getSelectedItem()),
				nullToEmptyString((String)form.getDesignationField().getSelectedItem()),
				nullToEmptyString((String)form.getDescriptionField().getSelectedItem()),
				convertToDoubleFromString(form.getAmountField().getText())
				);
	
		return s;
	}
	
	private Double convertToDoubleFromString(String s) {
		//try {
			return(Double.parseDouble(s));
		//}catch(Exception e) {
			//JOptionPane.showMessageDialog(null, message);
		//}
	}
	
	private void fillInDataUsingEnvelopeNumber(String s) {
		for(Donor d: form.getChurchDB()) {
			if (d.getEnvelopeNumber().compareToIgnoreCase(s)==0) {
				form.setFirstNameField(d.getFirstName());
				form.getLastNameField().setText(d.getLastName());
				form.setAddressField(d.getAddress());
				form.setCityField(d.getCity());
				form.setStateField(d.getState());
				form.setZipField(d.getZip());
				return;
			}
		}
			
	}
	
	private void showAllEntries(){
		
		JFrame frame = new JFrame();
		frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);

		Object columnNames[] = { "Last Name", "First Name", "Envelope Number", "Amount",
				"Designation","Category","Description","Address","City","State","Zip"};

		Object rowData[][] = new Object [offering.size()][columnNames.length];

		int i=0;
		for(Donation d: offering) {
			rowData[i][0] = d.getDonor().getLastName();
			rowData[i][1] = d.getDonor().getFirstName();
			rowData[i][2] = d.getDonor().getEnvelopeNumber();
			rowData[i][3] = d.getAmount();
			rowData[i][4] = d.getDesignation();
			rowData[i][5] = d.getCategory();
			rowData[i][6] = d.getDescription();
			rowData[i][7] = d.getDonor().getAddress();
			rowData[i][8] = d.getDonor().getCity();
			rowData[i][9] = d.getDonor().getState();
			rowData[i][10] = d.getDonor().getZip();

			i++;
		}

		JTable table = new JTable(rowData, columnNames);
		
		JScrollPane scrollPane = new JScrollPane(table);
		frame.add(scrollPane, BorderLayout.CENTER);
		frame.setSize(1200, 400);
		
		frame.setVisible(true);
	}

	private int checkForAddressMatchInDB() {
		int match = 0;
		for(Donor d: form.getChurchDB()) {
			if ((d.getLastName().compareToIgnoreCase(form.getLastNameField().getText())==0) && (d.getFirstName().compareToIgnoreCase(form.getFirstNameField().getText())==0)) {
				if((d.getEnvelopeNumber().compareToIgnoreCase(form.getEnvelopeField().getText())==0) &&
					   (d.getAddress().compareToIgnoreCase(form.getAddressField().getText())==0) &&
					   (d.getCity().compareToIgnoreCase(form.getCityField().getText())==0) &&
					   (d.getState().compareToIgnoreCase(form.getStateField().getSelectedItem().toString())==0) &&
					   (d.getZip().compareToIgnoreCase(form.getZipField().getText())==0)) {
					match=2;
					return match;
				}
			match=1;
			}
		}
		return match;
	}

	private void exportToExcel() {
		HSSFWorkbook wb = new HSSFWorkbook();
	
	    Sheet envSheet = wb.createSheet("Envelope");
	    Sheet plateSheet = wb.createSheet("Plate");
	    Sheet dfSheet = wb.createSheet("Designated Funds");
	    Sheet miscSheet = wb.createSheet("Misc");
	    Sheet dataSheet = wb.createSheet("Data");
	    Sheet totalsSheet = wb.createSheet("Totals");
	    Sheet treasurersReport = wb.createSheet("Treasurers Report");
	    Sheet checkReport = wb.createSheet("Check Amounts");
	    
	    CreationHelper createHelper = wb.getCreationHelper();
	    
	    createEnvSheet(wb, createHelper, envSheet);
	    createPlateSheet(wb, createHelper, plateSheet);
	    createDfSheet(wb, createHelper, dfSheet);
	    createMiscSheet(wb, createHelper, miscSheet);
	    createDataSheet(wb, createHelper, dataSheet);
	    createTotalsSheet(wb, createHelper, totalsSheet);
	    createTreasurersReport(wb, createHelper, treasurersReport);
	    createCheckReport(wb, createHelper, checkReport);
	    
	    try {
    			FileOutputStream fileOut = new FileOutputStream("workbook.xls");
    			wb.write(fileOut);
    			fileOut.close();
	    }catch(Exception e) {
	    	
	    }
	    
	    // Note that sheet name is Excel must not exceed 31 characters
	    // and must not contain any of the any of the following characters:
	    // 0x0000
	    // 0x0003
	    // colon (:)
	    // backslash (\)
	    // asterisk (*)
	    // question mark (?)
	    // forward slash (/)
	    // opening square bracket ([)
	    // closing square bracket (])

	    // You can use org.apache.poi.ss.util.WorkbookUtil#createSafeSheetName(String nameProposal)}
	    // for a safe way to create valid names, this utility replaces invalid characters with a space (' ')
	    // returns " O'Brien's sales   "

	    // Create a row and put some cells in it. Rows are 0 based.
	    //Row row = sheet1.createRow((short)0);
	    // Create a cell and put a value in it.
	    //Cell cell = row.createCell(0);
	    //cell.setCellValue(1);
	}
	
	private CellStyle topRowStyle(HSSFWorkbook wb) {
	    // setup background colors for sheet
	    HSSFPalette palette = wb.getCustomPalette();
	    HSSFColor myColor = palette.findSimilarColor(255, 202, 146);//(255, 202, 146)
	    short palIndex = myColor.getIndex();
		
	    // setup the cell style for titles on the page
	    CellStyle styleTopBar = wb.createCellStyle();
	    styleTopBar.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styleTopBar.setFillForegroundColor(palIndex);
	    styleTopBar.setAlignment(HorizontalAlignment.CENTER);
	    styleTopBar.setVerticalAlignment(VerticalAlignment.CENTER);
	    styleTopBar.setBorderBottom(BorderStyle.THIN);
	    //styleTopBar.setFont(font);
	    
	    return styleTopBar;
	}
	
    private void createEnvSheet(HSSFWorkbook wb, CreationHelper createHelper, Sheet envSheet) {
    	
		// setup column widths
	    envSheet.setColumnWidth(0, 3200);
	    envSheet.setColumnWidth(1, 3200);
	    envSheet.setColumnWidth(2, 3200);
	    envSheet.setColumnWidth(3, 3200);
	    envSheet.setColumnWidth(6, 3800);
		
	    // setup worksheet font
	    HSSFFont font= wb.createFont();
	    font.setFontHeightInPoints((short)10);
	    font.setFontName("Arial");
	    font.setColor(IndexedColors.BLACK.getIndex());
	    font.setBold(true);
	    font.setItalic(false);
	    
	    // setup background colors for sheet
	    HSSFPalette palette = wb.getCustomPalette();
	    HSSFColor myColor = palette.findSimilarColor(255, 202, 146);
	    short palIndex = myColor.getIndex();
	    
	    // setup the cell style for currency entries
	    CellStyle styleData = wb.createCellStyle();
	    styleData.setAlignment(HorizontalAlignment.RIGHT);
	    styleData.setDataFormat(wb.createDataFormat().getFormat(BuiltinFormats.getBuiltinFormat(8)));
	    
	    CellStyle envelopeStyle = wb.createCellStyle();
	    envelopeStyle.setAlignment(HorizontalAlignment.CENTER);
	    
	    //setup Styles for totals
	    CellStyle styleRightBorder = wb.createCellStyle();
	    styleRightBorder.setBorderRight(BorderStyle.THIN);
	    styleRightBorder.setAlignment(HorizontalAlignment.RIGHT);
	    styleRightBorder.setDataFormat(wb.createDataFormat().getFormat(BuiltinFormats.getBuiltinFormat(8)));
	    
	    CellStyle styleLeftBorder = wb.createCellStyle();
	    styleLeftBorder.setBorderLeft(BorderStyle.THIN);
	    
	    CellStyle styleBottomLeftCorner = wb.createCellStyle();
	    styleBottomLeftCorner.setBorderBottom(BorderStyle.THIN);
	    styleBottomLeftCorner.setBorderLeft(BorderStyle.THIN);
	    
	    CellStyle styleBottomLeftSum = wb.createCellStyle();
	    styleBottomLeftSum.setBorderBottom(BorderStyle.DOUBLE);
	    styleBottomLeftSum.setBorderLeft(BorderStyle.THIN);
	    
	    CellStyle styleBottomRightCorner = wb.createCellStyle();
	    styleBottomRightCorner.setBorderBottom(BorderStyle.THIN);
	    styleBottomRightCorner.setBorderRight(BorderStyle.THIN);
	    styleBottomRightCorner.setAlignment(HorizontalAlignment.RIGHT);
	    styleBottomRightCorner.setDataFormat(wb.createDataFormat().getFormat(BuiltinFormats.getBuiltinFormat(8)));
	    
	    CellStyle styleBottomRightSum = wb.createCellStyle();
	    styleBottomRightSum.setBorderBottom(BorderStyle.DOUBLE);
	    styleBottomRightSum.setBorderRight(BorderStyle.THIN);
	    styleBottomRightSum.setAlignment(HorizontalAlignment.RIGHT);
	    styleBottomRightSum.setDataFormat(wb.createDataFormat().getFormat(BuiltinFormats.getBuiltinFormat(8)));
	    
	    CellStyle styleTopLeftCorner = wb.createCellStyle();
	    styleTopLeftCorner.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styleTopLeftCorner.setFillForegroundColor(palIndex);
	    styleTopLeftCorner.setBorderTop(BorderStyle.THIN);
	    styleTopLeftCorner.setBorderLeft(BorderStyle.THIN);
	    
	    CellStyle styleTopRightCorner = wb.createCellStyle();
	    styleTopRightCorner.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styleTopRightCorner.setFillForegroundColor(palIndex);
	    styleTopRightCorner.setBorderTop(BorderStyle.THIN);
	    styleTopRightCorner.setBorderRight(BorderStyle.THIN);
	    
	    // create the titles
	    int r=0;
	    Row row = envSheet.createRow((short)r);
	    row.setHeightInPoints(20);
	    
	    Cell cell = row.createCell(0);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Envelope #"));

	    cell = row.createCell(1);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Check"));
	    
	    cell = row.createCell(2);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Cash"));
	    
	    cell = row.createCell(3);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("EFT PP"));
	    
	    for(Donation d: offering) {
	    		if(d.getDesignation().compareToIgnoreCase("Envelope - General Offering")==0){
	    			
	    			r++;
	    		    row = envSheet.createRow((short)r);
	    		    row.setHeightInPoints(18);
	    		    
	    		    cell= row.createCell(0);
	    		    cell.setCellStyle(envelopeStyle);
	    		    cell.setCellValue(createHelper.createRichTextString(d.getDonor().getEnvelopeNumber()));
	    		    
	    		    if(d.getCategory().compareToIgnoreCase("check")==0) {
	    		    		cell = row.createCell(1);
	    		    		cell.setCellStyle(styleData);
	    		    		cell.setCellValue(d.getAmount());
	    		    }
	    		    if(d.getCategory().compareToIgnoreCase("cash")==0) {
    		    			cell = row.createCell(2);
    		    			cell.setCellStyle(styleData);
    		    			cell.setCellValue(d.getAmount());
	    		    }
	    		    if(d.getCategory().compareToIgnoreCase("eft")==0) {
    		    			cell = row.createCell(3);
    		    			cell.setCellStyle(styleData);
    		    			cell.setCellValue(d.getAmount());
	    		    }
	    		}		
	    }
	    
	    if(envSheet.getRow((short)4)==null) {
			row = envSheet.createRow((short)4);
			row.setHeightInPoints(18);
			cell = row.createCell(5);
			cell.setCellStyle(styleTopLeftCorner);
			cell.setCellValue(createHelper.createRichTextString("TOTALS"));
			
			cell = envSheet.getRow((short)4).createCell(6);
			cell.setCellStyle(styleTopRightCorner);
			cell.setCellValue(createHelper.createRichTextString(" "));
	    }
	    else {
			cell = envSheet.getRow((short)4).createCell(5);
			cell.setCellStyle(styleTopLeftCorner);
			cell.setCellValue(createHelper.createRichTextString("TOTALS"));
			
			cell = envSheet.getRow((short)4).createCell(6);
			cell.setCellStyle(styleTopRightCorner);
			cell.setCellValue(createHelper.createRichTextString(" "));
	    }
	    
	    if(envSheet.getRow((short)5)==null) {
    			row = envSheet.createRow((short)5);
	    		row.setHeightInPoints(18);
	    		cell=row.createCell(5);
	    		cell.setCellStyle(styleLeftBorder);
	    		cell.setCellValue(createHelper.createRichTextString("CHECK"));
	    }	
	    else {
    			cell = envSheet.getRow((short)5).createCell(5);
    			cell.setCellStyle(styleLeftBorder);
    			cell.setCellValue(createHelper.createRichTextString("CHECK"));
	    }
	    if(envSheet.getRow((short)6)==null) {
			row=envSheet.createRow((short)6);
			row.setHeightInPoints(18);
    			cell=row.createCell(5);
    			cell.setCellStyle(styleLeftBorder);
    			cell.setCellValue(createHelper.createRichTextString("CASH"));
	    }
	    else {
			cell = envSheet.getRow((short)6).createCell(5);
			cell.setCellStyle(styleLeftBorder);
			cell.setCellValue(createHelper.createRichTextString("CASH"));
	    }
	    if(envSheet.getRow((short)7)==null) {
	    		row=envSheet.createRow((short)7);
	    		row.setHeightInPoints(18);
    			cell=row.createCell(5);
    			cell.setCellStyle(styleBottomLeftSum);
    			cell.setCellValue(createHelper.createRichTextString("EFT PP"));
	    }
	    else {
			cell = envSheet.getRow((short)7).createCell(5);
			cell.setCellStyle(styleBottomLeftSum);
			cell.setCellValue(createHelper.createRichTextString("EFT PP"));
	    }
	    if(envSheet.getRow((short)8)==null) {
    			row = envSheet.createRow((short)8);
    			row.setHeightInPoints(20);
    			cell=row.createCell(5);
    			cell.setCellStyle(styleBottomLeftCorner);
    			cell.setCellValue(createHelper.createRichTextString("ALL"));
	    }
	    else {
			cell = envSheet.getRow((short)8).createCell(5);
			cell.setCellStyle(styleBottomLeftCorner);
			cell.setCellValue(createHelper.createRichTextString("ALL"));
	    }
	    	    
	    row = envSheet.getRow(5);
		cell = row.createCell(6);
		cell.setCellStyle(styleRightBorder);
	    cell.setCellFormula("SUM(B:B)");
	    
	    row = envSheet.getRow(6);
		cell = row.createCell(6);
		cell.setCellStyle(styleRightBorder);
		cell.setCellFormula("SUM(C:C)");
		
		row = envSheet.getRow(7);
		cell = row.createCell(6);
		cell.setCellStyle(styleBottomRightSum);
		cell.setCellFormula("SUM(D:D)");
		
		row = envSheet.getRow(8);
		cell = row.createCell(6);
		cell.setCellStyle(styleBottomRightCorner);
		cell.setCellFormula("SUM(G6:G8)");    	
    }
 
    private void createPlateSheet(HSSFWorkbook wb, CreationHelper createHelper, Sheet plateSheet) {
    	
		// set the column widths for the sheet
		plateSheet.setColumnWidth(0, 5200);
		plateSheet.setColumnWidth(1, 5200);
		plateSheet.setColumnWidth(2, 3400);
		plateSheet.setColumnWidth(3, 3400);
		plateSheet.setColumnWidth(4, 3400);
		plateSheet.setColumnWidth(5, 6500);
		plateSheet.setColumnWidth(6, 3000);  // city
		plateSheet.setColumnWidth(7, 2000);  // state
		plateSheet.setColumnWidth(8, 3000);  // zip
		plateSheet.setColumnWidth(11, 3400); // $$ totals
		
		// set up font for the sheet
	    HSSFFont font= wb.createFont();
	    font.setFontHeightInPoints((short)10);
	    font.setFontName("Arial");
	    font.setColor(IndexedColors.BLACK.getIndex());
	    font.setBold(true);
	    font.setItalic(false);
	    
	    // set up the color for the top row
//	    HSSFPalette palette = wb.getCustomPalette();
//	    HSSFColor myColor = palette.findSimilarColor(255, 202, 146);
//	    short palIndex = myColor.getIndex();

	    
	    // setup the cell style for currency entries
	    CellStyle styleData = wb.createCellStyle();
	    styleData.setAlignment(HorizontalAlignment.RIGHT);
	    styleData.setDataFormat(wb.createDataFormat().getFormat( BuiltinFormats.getBuiltinFormat( 8 )));
	    
	    // create the titles
	    int r=0;
	    Row row = plateSheet.createRow((short)r);
	    row.setHeightInPoints(20);
	    
	    Cell cell = row.createCell(0);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("plate - First Name"));
	    
	    cell = row.createCell(1);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("plate - Last Name"));

	    cell = row.createCell(2);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Checks"));

	    cell = row.createCell(3);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Cash"));

	    cell = row.createCell(4);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("EFT PP"));

	    cell = row.createCell(5);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Address"));

	    cell = row.createCell(6);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("City"));

	    cell = row.createCell(7);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("State"));

	    cell = row.createCell(8);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Zip"));
	    
	    for(Donation d: offering) {
	    		if(d.getDesignation().compareToIgnoreCase("Plate - General Offering")==0) {
	    			
	    			r++;
	    		    row = plateSheet.createRow((short)r);
	    		    row.setHeightInPoints(18);

	    		    row.createCell(0).setCellValue(createHelper.createRichTextString(d.getDonor().getFirstName()));
	    		    row.createCell(1).setCellValue(createHelper.createRichTextString(d.getDonor().getLastName()));
	    		    if(d.getCategory().compareToIgnoreCase("check")==0) {
    		    			cell = row.createCell(2);
    		    			cell.setCellStyle(styleData);
    		    			cell.setCellValue(d.getAmount());
	    		    }
	    		    if(d.getCategory().compareToIgnoreCase("cash")==0) {
    		    			cell = row.createCell(3);
    		    			cell.setCellStyle(styleData);
    		    			cell.setCellValue(d.getAmount());
	    		    }
	    		    if(d.getCategory().compareToIgnoreCase("eft")==0) {
    		    			cell = row.createCell(4);
	    		    		cell.setCellStyle(styleData);
	    		    		cell.setCellValue(d.getAmount());
	    		    }
	    		    
	    		    row.createCell(5).setCellValue(createHelper.createRichTextString(d.getDonor().getAddress()));
	    		    row.createCell(6).setCellValue(createHelper.createRichTextString(d.getDonor().getCity()));
	    		    row.createCell(7).setCellValue(createHelper.createRichTextString(d.getDonor().getState()));
	    		    row.createCell(8).setCellValue(createHelper.createRichTextString(d.getDonor().getZip()));
	    		}		
	    }
	    
	    r++;
	    row = plateSheet.createRow((short)r);
	    //row.createCell(0).setCellValue(createHelper.createRichTextString(form.getDateLabel().getText()));
	   // row.createCell(1).setCellValue(createHelper.createRichTextString(" ");
	    row.createCell(0).setCellValue(createHelper.createRichTextString("unnamed plate cash"));
	    //row.createCell(3).setCellValue(createHelper.createRichTextString(d.getDonor().getLastName()));

	    	cell = row.createCell(3);
		cell.setCellStyle(styleData);
		cell.setCellFormula("Totals!E26");
	    
	    if(plateSheet.getRow((short)5)==null) {
    			row = plateSheet.createRow((short)5);
    			row.setHeightInPoints(18);
    			cell = row.createCell(10);
    			cell.setCellStyle(topRowStyle(wb));
    			cell.setCellValue(createHelper.createRichTextString("check"));
	    }
	    else	{
	    		row = plateSheet.getRow((short)5);
	    		cell = row.createCell(10);
	    		cell.setCellStyle(topRowStyle(wb));
	    		cell.setCellValue(createHelper.createRichTextString("check"));
	    }
	    if(plateSheet.getRow((short)6)==null) {
			row = plateSheet.createRow((short)6);
			row.setHeightInPoints(18);
			cell = row.createCell(10);
			cell.setCellStyle(topRowStyle(wb));
			cell.setCellValue(createHelper.createRichTextString("cash"));
	    }
	    else	{
	    		row = plateSheet.getRow((short)6);
    			cell = row.createCell(10);
    			cell.setCellStyle(topRowStyle(wb));
    			cell.setCellValue(createHelper.createRichTextString("cash"));
	    }
	    if(plateSheet.getRow((short)7)==null) {
			row = plateSheet.createRow((short)7);
			row.setHeightInPoints(18);
			cell = row.createCell(10);
			cell.setCellStyle(topRowStyle(wb));
			cell.setCellValue(createHelper.createRichTextString("EFT PP"));
	    }
	    else	{
	    		row = plateSheet.getRow((short)7);
    			cell = row.createCell(10);
    			cell.setCellStyle(topRowStyle(wb));
    			cell.setCellValue(createHelper.createRichTextString("EFT PP"));
	    }
	    if(plateSheet.getRow((short)8)==null) {
	 		row = plateSheet.createRow((short)8);
	 		row.setHeightInPoints(18);
	 		cell = row.createCell(10);
	 		cell.setCellStyle(topRowStyle(wb));
	 		cell.setCellValue(createHelper.createRichTextString("Total"));
	 	}
	    else	{
	 	    	row = plateSheet.getRow((short)8);
	     	cell = row.createCell(10);
	     	cell.setCellStyle(topRowStyle(wb));
	     	cell.setCellValue(createHelper.createRichTextString("Total"));
	    }
	    
	    // put in formula to compute the sum of all the CHECKS
	    row = plateSheet.getRow(5);
		cell = row.createCell(11);
		cell.setCellStyle(styleData);
	    cell.setCellFormula("SUM(C:C)");
	    
	    // put in formula to compute the sum of all the CASH
	    row = plateSheet.getRow(6);
		cell = row.createCell(11);
		cell.setCellStyle(styleData);
		cell.setCellFormula("SUM(D:D)");
		
	    // put in formula to compute the sum of all the EFTS
		row = plateSheet.getRow(7);
		cell = row.createCell(11);
		cell.setCellStyle(styleData);
		cell.setCellFormula("SUM(E:E)");
		
		// put in formula to compute the sum of all categories
		row = plateSheet.getRow(8);
		cell = row.createCell(11);
		cell.setCellStyle(styleData);
		cell.setCellFormula("SUM(L6:L8)");  
    }
  
    private void createDfSheet(HSSFWorkbook wb, CreationHelper createHelper, Sheet dfSheet) {

    		// setup column widths
    	    dfSheet.setColumnWidth(0, 2500);  // envelope
    	    dfSheet.setColumnWidth(1, 5200);	 // first name
    	    dfSheet.setColumnWidth(2, 4500);  // last name
    	    dfSheet.setColumnWidth(3, 3200);  // cash
    	    dfSheet.setColumnWidth(4, 3200);  // check
    	    dfSheet.setColumnWidth(5, 3200);  // eft
    	    dfSheet.setColumnWidth(6, 5200);  // fund name
    	    dfSheet.setColumnWidth(7, 5200);  // address
    	    dfSheet.setColumnWidth(8, 2500);  // city
    	    dfSheet.setColumnWidth(9, 2000);  // state
    	    dfSheet.setColumnWidth(10, 2500);  // zip
    	    dfSheet.setColumnWidth(13, 3400);  // money totals
    		
    	    // setup worksheet font
    	    HSSFFont font= wb.createFont();
    	    font.setFontHeightInPoints((short)10);
    	    font.setFontName("Arial");
    	    font.setColor(IndexedColors.BLACK.getIndex());
    	    font.setBold(true);
    	    font.setItalic(false);
    	    
    	    // setup the cell style for currency entries
    	    CellStyle styleData = wb.createCellStyle();
    	    styleData.setAlignment(HorizontalAlignment.RIGHT);
    	    styleData.setDataFormat(wb.createDataFormat().getFormat(BuiltinFormats.getBuiltinFormat(8)));
    	    
    	    // create the column titles
    	    int r=0;
    	    Row row = dfSheet.createRow((short)r);
    	    row.setHeightInPoints(20);
    	    
    	    Cell cell = row.createCell(0);
    	    cell.setCellStyle(topRowStyle(wb));
    	    cell.setCellValue(createHelper.createRichTextString("Envelope"));

    	    cell = row.createCell(1);
    	    cell.setCellStyle(topRowStyle(wb));
    	    cell.setCellValue(createHelper.createRichTextString("First Name"));
    	    
    	    cell = row.createCell(2);
    	    cell.setCellStyle(topRowStyle(wb));
    	    cell.setCellValue(createHelper.createRichTextString("Last Name"));
    	    
    	    cell = row.createCell(3);
    	    cell.setCellStyle(topRowStyle(wb));
    	    cell.setCellValue(createHelper.createRichTextString("Checks"));
    	    
    	    cell = row.createCell(4);
    	    cell.setCellStyle(topRowStyle(wb));
    	    cell.setCellValue(createHelper.createRichTextString("Cash"));

    	    cell = row.createCell(5);
    	    cell.setCellStyle(topRowStyle(wb));
    	    cell.setCellValue(createHelper.createRichTextString("EFT PP"));
    	    
    	    cell = row.createCell(6);
    	    cell.setCellStyle(topRowStyle(wb));
    	    cell.setCellValue(createHelper.createRichTextString("Fund"));
    	    
    	    cell = row.createCell(7);
    	    cell.setCellStyle(topRowStyle(wb));
    	    cell.setCellValue(createHelper.createRichTextString("Address"));
    	    
      	cell = row.createCell(8);
    	    cell.setCellStyle(topRowStyle(wb));
    	    cell.setCellValue(createHelper.createRichTextString("City"));

    	    cell = row.createCell(9);
    	    cell.setCellStyle(topRowStyle(wb));
    	    cell.setCellValue(createHelper.createRichTextString("State"));
    	    
    	    cell = row.createCell(10);
    	    cell.setCellStyle(topRowStyle(wb));
    	    cell.setCellValue(createHelper.createRichTextString("Zip"));
	    
	    for(Donation d: offering) {
	    		if(d.getDesignation().compareToIgnoreCase("Designated")==0) {
	    			
	    			r++;
	    		    row = dfSheet.createRow((short)r);
	    		    row.setHeightInPoints(18);

	    		    row.createCell(0).setCellValue(createHelper.createRichTextString(d.getDonor().getEnvelopeNumber()));
	    		    row.createCell(1).setCellValue(createHelper.createRichTextString(d.getDonor().getFirstName()));
	    		    row.createCell(2).setCellValue(createHelper.createRichTextString(d.getDonor().getLastName()));
	    		    if(d.getCategory().compareToIgnoreCase("check")==0) {
    		    			cell = row.createCell(3);
    		    			cell.setCellStyle(styleData);
    		    			cell.setCellValue(d.getAmount());
	    		    }
	    		    if(d.getCategory().compareToIgnoreCase("cash")==0) {
    		    			cell = row.createCell(4);
    		    			cell.setCellStyle(styleData);
    		    			cell.setCellValue(d.getAmount());
	    		    }
	    		    if(d.getCategory().compareToIgnoreCase("eft")==0) {
    		    			cell = row.createCell(5);
    		    			cell.setCellStyle(styleData);
    		    			cell.setCellValue(d.getAmount());
	    		    }
	    		    row.createCell(6).setCellValue(createHelper.createRichTextString(d.getDescription()));
	    		    row.createCell(7).setCellValue(createHelper.createRichTextString(d.getDonor().getAddress()));
	    		    row.createCell(8).setCellValue(createHelper.createRichTextString(d.getDonor().getCity()));
	    		    row.createCell(9).setCellValue(createHelper.createRichTextString(d.getDonor().getState()));
	    		    row.createCell(10).setCellValue(createHelper.createRichTextString(d.getDonor().getZip()));
	    		}		
	    }
	    
	    int col=12;
	    
	    if(dfSheet.getRow((short)5)==null) {
	    		row = dfSheet.createRow((short)5);
	    		row.setHeightInPoints(18);
	    		cell = row.createCell(col);
			cell.setCellStyle(topRowStyle(wb));
			cell.setCellValue(createHelper.createRichTextString("check"));
	    }
	    else {
    			row = dfSheet.getRow((short)5);
    			cell = row.createCell(col);
    			cell.setCellStyle(topRowStyle(wb));
    			cell.setCellValue(createHelper.createRichTextString("check"));
	    }
	    if(dfSheet.getRow((short)6)==null) {
    			row = dfSheet.createRow((short)6);
    			row.setHeightInPoints(18);
    			cell = row.createCell(col);
    			cell.setCellStyle(topRowStyle(wb));
    			cell.setCellValue(createHelper.createRichTextString("cash"));
	    }
	    else {
    			row = dfSheet.getRow((short)6);
    			cell = row.createCell(col);
    			cell.setCellStyle(topRowStyle(wb));
    			cell.setCellValue(createHelper.createRichTextString("cash"));
	    }
	    if(dfSheet.getRow((short)7)==null) {
			row = dfSheet.createRow((short)7);
			row.setHeightInPoints(18);
			cell = row.createCell(col);
			cell.setCellStyle(topRowStyle(wb));
			cell.setCellValue(createHelper.createRichTextString("EFT PP"));
	    }
	    else {
			row = dfSheet.getRow((short)7);
			cell = row.createCell(col);
			cell.setCellStyle(topRowStyle(wb));
			cell.setCellValue(createHelper.createRichTextString("EFT PP"));
	    }
	    if(dfSheet.getRow((short)8)==null) {
			row = dfSheet.createRow((short)8);
			row.setHeightInPoints(18);
			cell = row.createCell(col);
			cell.setCellStyle(topRowStyle(wb));
			cell.setCellValue(createHelper.createRichTextString("Total"));
	    }
	    else {
			row = dfSheet.getRow((short)8);
			cell = row.createCell(col);
			cell.setCellStyle(topRowStyle(wb));
			cell.setCellValue(createHelper.createRichTextString("Total"));
	    }

	    row = dfSheet.getRow(5);
		cell = row.createCell(col+1);
		cell.setCellStyle(styleData);
	    cell.setCellFormula("SUM(D:D)");
	    
	    row = dfSheet.getRow(6);
		cell = row.createCell(col+1);
		cell.setCellStyle(styleData);
		cell.setCellFormula("SUM(E:E)");
		
		row = dfSheet.getRow(7);
		cell = row.createCell(col+1);
		cell.setCellStyle(styleData);
		cell.setCellFormula("SUM(F:F)");
		
		row = dfSheet.getRow(8);
		cell = row.createCell(col+1);
		cell.setCellStyle(styleData);
		cell.setCellFormula("SUM(N6:N8)");    	
	    
    }
    
    private void createMiscSheet(HSSFWorkbook wb, CreationHelper createHelper, Sheet miscSheet) {

		// setup column widths
	    miscSheet.setColumnWidth(0, 2500);  // envelope
	    miscSheet.setColumnWidth(1, 5200);	 // first name
	    miscSheet.setColumnWidth(2, 4500);  // last name
	    miscSheet.setColumnWidth(3, 3200);  // cash
	    miscSheet.setColumnWidth(4, 3200);  // check
	    miscSheet.setColumnWidth(5, 3200);  // eft
	    miscSheet.setColumnWidth(6, 5200);  // fund name
	    miscSheet.setColumnWidth(7, 5200);  // address
	    miscSheet.setColumnWidth(8, 2500);  // city
	    miscSheet.setColumnWidth(9, 2000);  // state
	    miscSheet.setColumnWidth(10, 2500);  // zip
	    miscSheet.setColumnWidth(13, 3400);  // money totals
		
	    // setup worksheet font
	    HSSFFont font= wb.createFont();
	    font.setFontHeightInPoints((short)10);
	    font.setFontName("Arial");
	    font.setColor(IndexedColors.BLACK.getIndex());
	    font.setBold(true);
	    font.setItalic(false);
	    
	    // setup background colors for sheet
	    HSSFPalette palette = wb.getCustomPalette();
	    HSSFColor myColor = palette.findSimilarColor(255, 202, 146);
	    short palIndex = myColor.getIndex();
	    
	    // setup the cell style for currency entries
	    CellStyle styleData = wb.createCellStyle();
	    styleData.setAlignment(HorizontalAlignment.RIGHT);
	    styleData.setDataFormat(wb.createDataFormat().getFormat( BuiltinFormats.getBuiltinFormat(7)));
	    
	 // create the column titles
	    int r=0;
	    Row row = miscSheet.createRow((short)r);
	    row.setHeightInPoints(20);
	    
	    Cell cell = row.createCell(0);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Envelope"));

	    cell = row.createCell(1);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("First Name"));
	    
	    cell = row.createCell(2);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Last Name"));
	    
	    cell = row.createCell(3);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Checks"));
	    
	    cell = row.createCell(4);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Cash"));

	    cell = row.createCell(5);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("EFT PP"));
	    
	    cell = row.createCell(6);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Fund"));
	    
	    cell = row.createCell(7);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Address"));
	    
	    cell = row.createCell(8);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("City"));

	    cell = row.createCell(9);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("State"));
	    
	    cell = row.createCell(10);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Zip"));
	    
	    for(Donation d: offering) {
	    		if(d.getDesignation().compareToIgnoreCase("misc.")==0) {
	    			r++;
	    		    row = miscSheet.createRow((short)r);
	    		    row.createCell(0).setCellValue(createHelper.createRichTextString(d.getDonor().getEnvelopeNumber()));
	    		    row.createCell(1).setCellValue(createHelper.createRichTextString(d.getDonor().getFirstName()));
	    		    row.createCell(2).setCellValue(createHelper.createRichTextString(d.getDonor().getLastName()));
	    		    if(d.getCategory().compareToIgnoreCase("check")==0) {
	    		    		cell = row.createCell(3);
	    		    		cell.setCellStyle(styleData);
	    		    		cell.setCellValue(d.getAmount());
	    		    }
	    		    if(d.getCategory().compareToIgnoreCase("cash")==0) {
    		    			cell = row.createCell(4);
    		    			cell.setCellStyle(styleData);
    		    			cell.setCellValue(d.getAmount());
	    		    }
	    		    if(d.getCategory().compareToIgnoreCase("eft")==0) {
    		    			cell = row.createCell(5);
    		    			cell.setCellStyle(styleData);
    		    			cell.setCellValue(d.getAmount());
	    		    }
	    		    row.createCell(6).setCellValue(createHelper.createRichTextString(d.getDescription()));
	    		    row.createCell(7).setCellValue(createHelper.createRichTextString(d.getDonor().getAddress()));
	    		    row.createCell(8).setCellValue(createHelper.createRichTextString(d.getDonor().getCity()));
	    		    row.createCell(9).setCellValue(createHelper.createRichTextString(d.getDonor().getState()));
	    		    row.createCell(10).setCellValue(createHelper.createRichTextString(d.getDonor().getZip()));
	    		}		
	    }
	    
	    int col=12;
	    
	    /*
	    if(miscSheet.getRow((short)5)==null)
			miscSheet.createRow((short)5).createCell(col).setCellValue(createHelper.createRichTextString("check"));
	    else
			miscSheet.getRow((short)5).createCell(col).setCellValue(createHelper.createRichTextString("check"));

	    if(miscSheet.getRow((short)6)==null)
	    		miscSheet.createRow((short)6).createCell(col).setCellValue(createHelper.createRichTextString("cash"));
	    else
	    		miscSheet.getRow((short)6).createCell(col).setCellValue(createHelper.createRichTextString("cash"));

	    if(miscSheet.getRow((short)7)==null)
    			miscSheet.createRow((short)7).createCell(col).setCellValue(createHelper.createRichTextString("EFT PP"));
	    else
    			miscSheet.getRow((short)7).createCell(col).setCellValue(createHelper.createRichTextString("EFT PP"));
*/
	    //miscSheet.getRow(5).createCell(col+1).setCellValue(check);
	    //miscSheet.getRow(6).createCell(col+1).setCellValue(cash);
	    //miscSheet.getRow(7).createCell(col+1).setCellValue(EFT);
	    
	    
	    if(miscSheet.getRow((short)5)==null) {
	    		row = miscSheet.createRow((short)5);
	    		row.setHeightInPoints(18);
	    		cell = row.createCell(col);
			cell.setCellStyle(topRowStyle(wb));
			cell.setCellValue(createHelper.createRichTextString("check"));
	    }
	    else {
    			row = miscSheet.getRow((short)5);
    			cell = row.createCell(col);
    			cell.setCellStyle(topRowStyle(wb));
    			cell.setCellValue(createHelper.createRichTextString("check"));
	    }
	    if(miscSheet.getRow((short)6)==null) {
    			row = miscSheet.createRow((short)6);
    			row.setHeightInPoints(18);
    			cell = row.createCell(col);
    			cell.setCellStyle(topRowStyle(wb));
    			cell.setCellValue(createHelper.createRichTextString("cash"));
	    }
	    else {
    			row = miscSheet.getRow((short)6);
    			cell = row.createCell(col);
    			cell.setCellStyle(topRowStyle(wb));
    			cell.setCellValue(createHelper.createRichTextString("cash"));
	    }
	    if(miscSheet.getRow((short)7)==null) {
			row = miscSheet.createRow((short)7);
			row.setHeightInPoints(18);
			cell = row.createCell(col);
			cell.setCellStyle(topRowStyle(wb));
			cell.setCellValue(createHelper.createRichTextString("EFT PP"));
	    }
	    else {
			row = miscSheet.getRow((short)7);
			cell = row.createCell(col);
			cell.setCellStyle(topRowStyle(wb));
			cell.setCellValue(createHelper.createRichTextString("EFT PP"));
	    }
	    if(miscSheet.getRow((short)8)==null) {
			row = miscSheet.createRow((short)8);
			row.setHeightInPoints(18);
			cell = row.createCell(col);
			cell.setCellStyle(topRowStyle(wb));
			cell.setCellValue(createHelper.createRichTextString("Total"));
	    }
	    else {
			row = miscSheet.getRow((short)8);
			cell = row.createCell(col);
			cell.setCellStyle(topRowStyle(wb));
			cell.setCellValue(createHelper.createRichTextString("Total"));
	    }

	    row = miscSheet.getRow(5);
		cell = row.createCell(col+1);
		cell.setCellStyle(styleData);
	    cell.setCellFormula("SUM(D:D)");
	    
	    row = miscSheet.getRow(6);
		cell = row.createCell(col+1);
		cell.setCellStyle(styleData);
		cell.setCellFormula("SUM(E:E)");
		
		row = miscSheet.getRow(7);
		cell = row.createCell(col+1);
		cell.setCellStyle(styleData);
		cell.setCellFormula("SUM(F:F)");
		
		row = miscSheet.getRow(8);
		cell = row.createCell(col+1);
		cell.setCellStyle(styleData);
		cell.setCellFormula("SUM(N6:N8)");   
	    
    }
  
    private void createDataSheet(HSSFWorkbook wb, CreationHelper createHelper, Sheet dataSheet) {
    	
	    // setup column widths
    		dataSheet.setColumnWidth(0, 3000);  // date
    		dataSheet.setColumnWidth(1, 2000);  // envelope #
    		dataSheet.setColumnWidth(2, 5400);  // first name
    		dataSheet.setColumnWidth(3, 4400);  // last name
    		dataSheet.setColumnWidth(4, 3000);  // check
    		dataSheet.setColumnWidth(5, 3000);  // cash
    		dataSheet.setColumnWidth(6, 3000);  // eft
    		dataSheet.setColumnWidth(7, 5800);  // category 
    		dataSheet.setColumnWidth(8, 5200);  // description
    		dataSheet.setColumnWidth(9, 5400);  // address
    		dataSheet.setColumnWidth(10, 3400); // city
    		dataSheet.setColumnWidth(11, 2500); // state  
    		dataSheet.setColumnWidth(11, 3000); // zip
		
	    HSSFFont font= wb.createFont();
	    font.setFontHeightInPoints((short)10);
	    font.setFontName("Arial");
	    font.setColor(IndexedColors.BLACK.getIndex());
	    font.setBold(true);
	    font.setItalic(false);
	    
	    // setup background colors for sheet
	    HSSFPalette palette = wb.getCustomPalette();
	    HSSFColor myColor = palette.findSimilarColor(255, 202, 146);
	    short palIndex = myColor.getIndex();
	    
	    // setup the cell style for currency entries
	    CellStyle styleData = wb.createCellStyle();
	    styleData.setAlignment(HorizontalAlignment.RIGHT);
	    styleData.setDataFormat(wb.createDataFormat().getFormat(BuiltinFormats.getBuiltinFormat(7)));
	    
	    float check=0; 
	    float cash=0;
	    float EFT=0;

	    int r=0;
	    Row row = dataSheet.createRow((short)r);
	    row.setHeightInPoints(20);
	    
	    Cell cell = row.createCell(0);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Date"));
	    
	    cell = row.createCell(1);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Envelope"));
	    
	    cell = row.createCell(2);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("First Name"));
	    
	    cell = row.createCell(3);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Last Name"));
	    
	    cell = row.createCell(4);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Checks"));
	    
	    cell = row.createCell(5);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Cash"));
	    
	    cell = row.createCell(6);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("EFT PP"));
	    
	    cell = row.createCell(7);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Category"));
	    
	    cell = row.createCell(8);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Description"));
	    
	    cell = row.createCell(9);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Address"));
	    
	    cell = row.createCell(10);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("City"));
	    
	    cell = row.createCell(11);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("State"));
	    
	    cell = row.createCell(12);
	    cell.setCellStyle(topRowStyle(wb));
	    cell.setCellValue(createHelper.createRichTextString("Zip"));
	    
	    for(Donation d: offering) {
	    			r++;
	    		    row = dataSheet.createRow((short)r);
	    		    row.createCell(0).setCellValue(createHelper.createRichTextString(form.getDateLabel().getText()));
	    		    row.createCell(1).setCellValue(createHelper.createRichTextString(d.getDonor().getEnvelopeNumber()));
	    		    row.createCell(2).setCellValue(createHelper.createRichTextString(d.getDonor().getFirstName()));
	    		    row.createCell(3).setCellValue(createHelper.createRichTextString(d.getDonor().getLastName()));
	    		    if(d.getCategory().compareToIgnoreCase("check")==0) {
		    			cell = row.createCell(4);
		    			cell.setCellStyle(styleData);
		    			cell.setCellValue(d.getAmount());
	    		    		check+=d.getAmount();
	    		    }
	    		    if(d.getCategory().compareToIgnoreCase("cash")==0) {
	    		    		cell = row.createCell(5);
		    			cell.setCellStyle(styleData);
		    			cell.setCellValue(d.getAmount());
    		    			cash+=d.getAmount();
	    		    }
	    		    if(d.getCategory().compareToIgnoreCase("eft")==0) {cell = row.createCell(4);
		    			cell = row.createCell(6);
		    			cell.setCellStyle(styleData);
		    			cell.setCellValue(d.getAmount());
    		    			EFT+=d.getAmount();
	    		    }
	    		    row.createCell(7).setCellValue(createHelper.createRichTextString(d.getDesignation()));
	    		    row.createCell(8).setCellValue(createHelper.createRichTextString(d.getDescription()));
	    		    row.createCell(9).setCellValue(createHelper.createRichTextString(d.getDonor().getAddress()));
	    		    row.createCell(10).setCellValue(createHelper.createRichTextString(d.getDonor().getCity()));
	    		    row.createCell(11).setCellValue(createHelper.createRichTextString(d.getDonor().getState()));
	    		    row.createCell(12).setCellValue(createHelper.createRichTextString(d.getDonor().getZip()));		
	    }
	    
	    r++;
	    row = dataSheet.createRow((short)r);
	    row.createCell(0).setCellValue(createHelper.createRichTextString(form.getDateLabel().getText()));
	   // row.createCell(1).setCellValue(createHelper.createRichTextString(" ");
	    row.createCell(2).setCellValue(createHelper.createRichTextString("Unnamed Plate Cash"));
	    //row.createCell(3).setCellValue(createHelper.createRichTextString(d.getDonor().getLastName()));
	    row.createCell(7).setCellValue(createHelper.createRichTextString("Plate - General Offering"));

	    	cell = row.createCell(5);
		cell.setCellStyle(styleData);
		cell.setCellFormula("Totals!E26");
    		//cash+=d.getAmount();
    		
    		
	    /*
	    int col=14;
	    
	    if(dataSheet.getRow((short)5)==null)
			dataSheet.createRow((short)5).createCell(col).setCellValue(createHelper.createRichTextString("check"));
	    else
			dataSheet.getRow((short)5).createCell(col).setCellValue(createHelper.createRichTextString("check"));

	    if(dataSheet.getRow((short)6)==null)
	    		dataSheet.createRow((short)6).createCell(col).setCellValue(createHelper.createRichTextString("cash"));
	    else
	    		dataSheet.getRow((short)6).createCell(col).setCellValue(createHelper.createRichTextString("cash"));

	    if(dataSheet.getRow((short)7)==null)
    			dataSheet.createRow((short)7).createCell(col).setCellValue(createHelper.createRichTextString("EFT PP"));
	    else
    			dataSheet.getRow((short)7).createCell(col).setCellValue(createHelper.createRichTextString("EFT PP"));

		cell = dataSheet.getRow(5).createCell(col+1);
		cell.setCellStyle(styleData);
		cell.setCellValue(check);
		
		cell = dataSheet.getRow(6).createCell(col+1);
		cell.setCellStyle(styleData);
		cell.setCellValue(cash);
		
		cell = dataSheet.getRow(7).createCell(col+1);
		cell.setCellStyle(styleData);
		cell.setCellValue(EFT);
		*/
    }
    
    private void createTotalsSheet(HSSFWorkbook wb, CreationHelper createHelper, Sheet sheet) {
    		Row row;
    		Cell cell;
    		
    		// setup column widths
	    sheet.setColumnWidth(0, 3200);
		
	    // setup worksheet font
	    HSSFFont font= wb.createFont();
	    font.setFontHeightInPoints((short)10);
	    font.setFontName("Arial");
	    font.setColor(IndexedColors.BLACK.getIndex());
	    font.setBold(true);
	    font.setItalic(false);
	    
	    // setup background colors for sheet
	    HSSFPalette palette = wb.getCustomPalette();
	    HSSFColor myColor = palette.findSimilarColor(255, 202, 146);
	    short palIndex = myColor.getIndex();
	    
	    // create the necessary rows
	    for(int r = 0; r<30 ; r++) {
	    		row = sheet.createRow((short)r);
	    	    row.setHeightInPoints(20);
	    }
	    
	    ////////// styles  //////////////////////////////////////////////////////
		
		CellStyle toprowmerged = wb.createCellStyle();
		toprowmerged.setAlignment(HorizontalAlignment.CENTER);
		toprowmerged.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    toprowmerged.setFillForegroundColor(palIndex);
	    toprowmerged.setBorderTop(BorderStyle.THIN);
	    toprowmerged.setBorderRight(BorderStyle.THIN);
	    toprowmerged.setBorderLeft(BorderStyle.THIN);
	    toprowmerged.setFont(font);
	    
	    CellStyle coloredstyle = wb.createCellStyle();
	    coloredstyle.setAlignment(HorizontalAlignment.CENTER);
	    coloredstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    coloredstyle.setFillForegroundColor(palIndex);
	    coloredstyle.setFont(font);
	    
	    CellStyle uncoloredstyle = wb.createCellStyle();
	    uncoloredstyle.setAlignment(HorizontalAlignment.RIGHT);
	    uncoloredstyle.setDataFormat(wb.createDataFormat().getFormat( BuiltinFormats.getBuiltinFormat(7)));
	    
	    CellStyle coloredtoprow=wb.createCellStyle();
	    coloredtoprow.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    coloredtoprow.setFillForegroundColor(palIndex);
	    coloredtoprow.setBorderTop(BorderStyle.THIN);
	    
	    CellStyle uncoloredtoprow=wb.createCellStyle();
	    uncoloredtoprow.setBorderTop(BorderStyle.THIN);
	    uncoloredtoprow.setDataFormat(wb.createDataFormat().getFormat( BuiltinFormats.getBuiltinFormat(7)));
	    
	    CellStyle coloredbottomrow=wb.createCellStyle();
	    coloredbottomrow.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    coloredbottomrow.setFillForegroundColor(palIndex);
	    coloredbottomrow.setBorderBottom(BorderStyle.THIN);
	    
	    CellStyle uncoloredbottomrow=wb.createCellStyle();
	    uncoloredbottomrow.setBorderBottom(BorderStyle.THIN);
	    uncoloredbottomrow.setDataFormat(wb.createDataFormat().getFormat( BuiltinFormats.getBuiltinFormat(7)));
	    
	    CellStyle uncoloredbottomrowwithtopdouble=wb.createCellStyle();
	    uncoloredbottomrowwithtopdouble.setBorderBottom(BorderStyle.THIN);
	    uncoloredbottomrowwithtopdouble.setBorderTop(BorderStyle.DOUBLE);
	    uncoloredbottomrowwithtopdouble.setDataFormat(wb.createDataFormat().getFormat( BuiltinFormats.getBuiltinFormat(7)));
	    
	    CellStyle coloredleftcolumn=wb.createCellStyle();
	    coloredleftcolumn.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    coloredleftcolumn .setFillForegroundColor(palIndex);
	    coloredleftcolumn.setFont(font);
	    coloredleftcolumn.setBorderLeft(BorderStyle.THIN);
	    
	    CellStyle uncoloredleftcolumn=wb.createCellStyle();
	    uncoloredleftcolumn.setBorderLeft(BorderStyle.THIN);
	    uncoloredleftcolumn.setDataFormat(wb.createDataFormat().getFormat( BuiltinFormats.getBuiltinFormat(7)));
	    
	    CellStyle coloredrightcolumn = wb.createCellStyle();
	    coloredrightcolumn.setAlignment(HorizontalAlignment.CENTER);
	    coloredrightcolumn.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    coloredrightcolumn.setFillForegroundColor(palIndex);
	    coloredrightcolumn.setFont(font);
	    coloredrightcolumn.setBorderRight(BorderStyle.THIN);
	    
	    CellStyle uncoloredrightcolumn=wb.createCellStyle();
	    uncoloredrightcolumn.setBorderRight(BorderStyle.THIN);
	    uncoloredrightcolumn.setAlignment(HorizontalAlignment.RIGHT);
	    uncoloredrightcolumn.setDataFormat(wb.createDataFormat().getFormat( BuiltinFormats.getBuiltinFormat(7)));
	    
	    CellStyle coloredtoprightcorner = wb.createCellStyle();
	    coloredtoprightcorner.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    coloredtoprightcorner.setFillForegroundColor(palIndex);
	    coloredtoprightcorner.setBorderTop(BorderStyle.THIN);
	    coloredtoprightcorner.setBorderRight(BorderStyle.THIN);
	    
	    CellStyle coloredtopleftcorner = wb.createCellStyle();
	    coloredtopleftcorner.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    coloredtopleftcorner.setFillForegroundColor(palIndex);
	    coloredtopleftcorner.setFont(font);
	    coloredtopleftcorner.setBorderTop(BorderStyle.THIN);
	    coloredtopleftcorner.setBorderLeft(BorderStyle.THIN);
	    
	    CellStyle coloredbottomrightcorner = wb.createCellStyle();
	    coloredbottomrightcorner.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    coloredbottomrightcorner.setFillForegroundColor(palIndex);
	    coloredbottomrightcorner.setBorderBottom(BorderStyle.THIN);
	    coloredbottomrightcorner.setBorderRight(BorderStyle.THIN);
	    
	    CellStyle uncoloredbottomrightcorner = wb.createCellStyle();
	    uncoloredbottomrightcorner.setBorderBottom(BorderStyle.THIN);
	    uncoloredbottomrightcorner.setBorderRight(BorderStyle.THIN);
	    uncoloredbottomrightcorner.setDataFormat(wb.createDataFormat().getFormat( BuiltinFormats.getBuiltinFormat(7)));
	    
	    CellStyle uncoloredbottomrightcornerwithtopdouble = wb.createCellStyle();
	    uncoloredbottomrightcornerwithtopdouble.setBorderBottom(BorderStyle.THIN);
	    uncoloredbottomrightcornerwithtopdouble.setBorderRight(BorderStyle.THIN);
	    uncoloredbottomrightcornerwithtopdouble.setBorderTop(BorderStyle.DOUBLE);
	    uncoloredbottomrightcornerwithtopdouble.setDataFormat(wb.createDataFormat().getFormat( BuiltinFormats.getBuiltinFormat(7)));
	    
	    CellStyle coloredbottomleftcorner = wb.createCellStyle();
	    coloredbottomleftcorner.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    coloredbottomleftcorner.setFillForegroundColor(palIndex);
	    coloredbottomleftcorner.setBorderBottom(BorderStyle.THIN);
	    coloredbottomleftcorner.setBorderLeft(BorderStyle.THIN);
	    coloredbottomleftcorner.setFont(font);
	    
	    CellStyle coloredbottomleftcornerwithtopdoubleborder = wb.createCellStyle();
	    coloredbottomleftcornerwithtopdoubleborder.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    coloredbottomleftcornerwithtopdoubleborder.setFillForegroundColor(palIndex);
	    coloredbottomleftcornerwithtopdoubleborder.setBorderBottom(BorderStyle.THIN);
	    coloredbottomleftcornerwithtopdoubleborder.setBorderLeft(BorderStyle.THIN);
	    coloredbottomleftcornerwithtopdoubleborder.setBorderTop(BorderStyle.DOUBLE);
	    coloredbottomleftcornerwithtopdoubleborder.setFont(font);
	
	    ///////////////setting the column widths//////////////////////////////////////////////////////////////

	    sheet.setColumnWidth(2, 3900);
	    sheet.setColumnWidth(3, 3900);
	    sheet.setColumnWidth(4, 3900);
	    sheet.setColumnWidth(5, 3900);
	    sheet.setColumnWidth(6, 3900);
	    sheet.setColumnWidth(7, 3900);
	
	    //////////////////////////////////////////////////////////////////////////////////////////////////////
	    
	    row = sheet.getRow((short)2);
	    
		row = sheet.getRow((short)2);
	    cell = row.createCell(3);
	    cell.setCellStyle(coloredtoprow);
		cell.setCellValue(" ");
		
		row = sheet.getRow((short)2);
	    cell = row.createCell(4);
	    cell.setCellStyle(coloredtoprow);
		cell.setCellValue(" ");
		
		row = sheet.getRow((short)2);
	    cell = row.createCell(5);
	    cell.setCellStyle(coloredtoprow);
		cell.setCellValue(" ");
	    
		row = sheet.getRow((short)2);
	    cell = row.createCell(6);
	    cell.setCellStyle(coloredtoprow);
		cell.setCellValue(" ");
		
		row = sheet.getRow((short)2);
	    cell = row.createCell(7);
	    cell.setCellStyle(coloredtoprightcorner);
		cell.setCellValue(" ");
		
	    CellRangeAddress cellRangeAddress = new CellRangeAddress(2, 2, 2,7);
	    sheet.addMergedRegion(cellRangeAddress);
	    cell = row.createCell(2);
	    cell.setCellStyle(toprowmerged);
		cell.setCellValue("WEEKLY TOTALS");
		
		row = sheet.getRow((short)3);
	    cell = row.createCell(2);
	    cell.setCellStyle(coloredleftcolumn);
		cell.setCellValue(" ");
		
	    row = sheet.getRow((short)3);
	    cell = row.createCell(3);
	    cell.setCellStyle(coloredstyle);
		cell.setCellValue("Checks");

		row = sheet.getRow((short)3);
	    cell = row.createCell(4);
	    cell.setCellStyle(coloredstyle);
		cell.setCellValue("Cash");
		
		row = sheet.getRow((short)3);
	    cell = row.createCell(5);
	    cell.setCellStyle(coloredstyle);
		cell.setCellValue("Total (Bank DPST)");
		
		row = sheet.getRow((short)3);
	    cell = row.createCell(6);
	    cell.setCellStyle(coloredstyle);
		cell.setCellValue("EFT PP");
		
		row = sheet.getRow((short)3);
	    cell = row.createCell(7);
	    cell.setCellStyle(coloredrightcolumn);
		cell.setCellValue("Week Total");
		
		row = sheet.getRow((short)4);
	    cell = row.createCell(2);
	    cell.setCellStyle(coloredleftcolumn);
		cell.setCellValue("Misc.");
		
		// sum of the misc. checks
		row = sheet.getRow((short)4);
	    cell = row.createCell(3);
	    cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUMIF(Data!$H:$H,\"Misc.\",Data!E:E)");
		
		// sum of misc. cash
		row = sheet.getRow((short)4);
	    cell = row.createCell(4);
	    cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUMIF(Data!$H:$H,\"Misc.\",Data!F:F)");
		
		// total bank deposit for misc.
		row = sheet.getRow((short)4);
	    cell = row.createCell(5);
	    cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUM(D5:E5)");
		
		// sum of misc. EFTs
		row = sheet.getRow((short)4);
	    cell = row.createCell(6);
	    cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUMIF(Data!$H:$H,\"Misc.\",Data!G:G)");
		
		// sum of all misc.
		row = sheet.getRow((short)4);
	    cell = row.createCell(7);
	    cell.setCellStyle(uncoloredrightcolumn);
		cell.setCellFormula("F5+G5");
		
		row = sheet.getRow((short)5);
	    cell = row.createCell(2);
	    cell.setCellStyle(coloredleftcolumn);
		cell.setCellValue("Plate");
		
		row = sheet.getRow((short)5);
	    cell = row.createCell(3);
	    cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUMIF(Data!$H:$H,\"Plate - General Offering\",Data!E:E)");
		
		row = sheet.getRow((short)5);
	    cell = row.createCell(4);
	    cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUMIF(Data!$H:$H,\"Plate - General Offering\",Data!F:F)");
		
		row = sheet.getRow((short)5);
	    cell = row.createCell(5);
	    cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUM(D6:E6)");
		
		row = sheet.getRow((short)5);
	    cell = row.createCell(6);
	    cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUMIF(Data!$H:$H,\"Plate - General Offering\",Data!G:G)");
		
		row = sheet.getRow((short)5);
	    cell = row.createCell(7);
	    cell.setCellStyle(uncoloredrightcolumn);
		cell.setCellFormula("F6+G6");
		
		row = sheet.getRow((short)6);
	    cell = row.createCell(2);
	    cell.setCellStyle(coloredleftcolumn);
		cell.setCellValue("Designated");
		
		row = sheet.getRow((short)6);
	    cell = row.createCell(3);
	    cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUMIF(Data!$H:$H,\"Designated\",Data!E:E)");
		
		row = sheet.getRow((short)6);
	    cell = row.createCell(4);
	    cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUMIF(Data!$H:$H,\"Designated\",Data!F:F)");
		
		row = sheet.getRow((short)6);
	    cell = row.createCell(5);
	    cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUM(D7:E7)");
		
		row = sheet.getRow((short)6);
	    cell = row.createCell(6);
	    cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUMIF(Data!$H:$H,\"Designated\",Data!G:G)");
		
		row = sheet.getRow((short)6);
	    cell = row.createCell(7);
	    cell.setCellStyle(uncoloredrightcolumn);
		cell.setCellFormula("F7+G7");
		
		row = sheet.getRow((short)7);
	    cell = row.createCell(2);
	    cell.setCellStyle(coloredleftcolumn);
		cell.setCellValue("Envelope");
		
		row = sheet.getRow((short)7);
	    cell = row.createCell(3);
	    cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUMIF(Data!$H:$H,\"Envelope - General Offering\",Data!E:E)");
		
		row = sheet.getRow((short)7);
	    cell = row.createCell(4);
	    cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUMIF(Data!$H:$H,\"Envelope - General Offering\",Data!F:F)");
		
		row = sheet.getRow((short)7);
	    cell = row.createCell(5);
	    cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUM(D8:E8)");
		
		row = sheet.getRow((short)7);
	    cell = row.createCell(6);
	    cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUMIF(Data!$H:$H,\"Envelope - General Offering\",Data!G:G)");
		
		row = sheet.getRow((short)7);
	    cell = row.createCell(7);
	    cell.setCellStyle(uncoloredrightcolumn);
		cell.setCellFormula("F8+G8");
		
		row = sheet.getRow((short)8);
	    cell = row.createCell(2);
	    cell.setCellStyle(coloredbottomleftcornerwithtopdoubleborder);
		cell.setCellValue("Grand Totals");
		
		row = sheet.getRow((short)8);
	    cell = row.createCell(3);
	    cell.setCellStyle(uncoloredbottomrowwithtopdouble);
		cell.setCellFormula("SUM(D5:D8)");
		
		row = sheet.getRow((short)8);
	    cell = row.createCell(4);
	    cell.setCellStyle(uncoloredbottomrowwithtopdouble);
		cell.setCellFormula("SUM(E5:E8)");
		
		row = sheet.getRow((short)8);
	    cell = row.createCell(5);
	    cell.setCellStyle(uncoloredbottomrowwithtopdouble);
		cell.setCellFormula("SUM(F5:F8)");
		
		row = sheet.getRow((short)8);
	    cell = row.createCell(6);
	    cell.setCellStyle(uncoloredbottomrowwithtopdouble);
		cell.setCellFormula("SUM(G5:G8)");

		row = sheet.getRow((short)8);
	    cell = row.createCell(7);
	    cell.setCellStyle(uncoloredbottomrightcornerwithtopdouble);
		cell.setCellFormula("F9+G9");
		
		row=sheet.getRow(4);
	    CellRangeAddress cellRangeAddress2 = new CellRangeAddress(4, 4, 10,11);
	    sheet.addMergedRegion(cellRangeAddress2);
	    cell = row.createCell(10);
	    cell.setCellStyle(toprowmerged);
		cell.setCellValue("BANK DEPOSIT SLIP");
		
		row = sheet.getRow((short)4);
	    cell = row.createCell(11);
	    cell.setCellStyle(coloredtoprightcorner);
		cell.setCellValue(" ");

		row = sheet.getRow((short)5);
	    cell = row.createCell(10);
	    cell.setCellStyle(coloredleftcolumn);
		cell.setCellValue("Date");
		
		row = sheet.getRow((short)5);
	    cell = row.createCell(11);
	    cell.setCellStyle(uncoloredrightcolumn);
		cell.setCellValue(" ");

		row = sheet.getRow((short)6);
	    cell = row.createCell(10);
	    cell.setCellStyle(coloredleftcolumn);
		cell.setCellValue("Coins");
		
		row = sheet.getRow((short)6);
	    cell = row.createCell(11);
	    cell.setCellStyle(uncoloredrightcolumn);
		cell.setCellValue(" ");
		
		row = sheet.getRow((short)7);
	    cell = row.createCell(10);
	    cell.setCellStyle(coloredleftcolumn);
		cell.setCellValue("Bills");
		
		row = sheet.getRow((short)7);
	    cell = row.createCell(11);
	    cell.setCellStyle(uncoloredrightcolumn);
		cell.setCellValue(" ");

		row = sheet.getRow((short)8);
	    cell = row.createCell(10);
	    cell.setCellStyle(coloredleftcolumn);
		cell.setCellValue("Checks");
		
		row = sheet.getRow((short)8);
	    cell = row.createCell(11);
	    cell.setCellStyle(uncoloredrightcolumn);
		cell.setCellValue(" ");
		
		row = sheet.getRow((short)9);
	    cell = row.createCell(10);
	    cell.setCellStyle(coloredbottomleftcornerwithtopdoubleborder);
		cell.setCellValue("Total");
		
		row = sheet.getRow((short)9);
	    cell = row.createCell(11);
	    cell.setCellStyle(uncoloredbottomrightcornerwithtopdouble);
		cell.setCellValue(" ");

		// Cash reconciliation section
		row=sheet.getRow(11);
	    CellRangeAddress cellRangeAddress3 = new CellRangeAddress(11, 11, 2, 8);
	    sheet.addMergedRegion(cellRangeAddress3);
	    cell = row.createCell(2);
	    cell.setCellStyle(toprowmerged);
		cell.setCellValue("CASH RECONCILIATION");
		
		row = sheet.getRow((short)11);
	    cell = row.createCell(3);
	    cell.setCellStyle(coloredtoprow);
		cell.setCellValue(" ");
		
		row = sheet.getRow((short)11);
	    cell = row.createCell(4);
	    cell.setCellStyle(coloredtoprow);
		cell.setCellValue(" ");
		
		row = sheet.getRow((short)11);
	    cell = row.createCell(5);
	    cell.setCellStyle(coloredtoprow);
		cell.setCellValue(" ");
		
		row = sheet.getRow((short)11);
	    cell = row.createCell(6);
	    cell.setCellStyle(coloredtoprow);
		cell.setCellValue(" ");
		
		row = sheet.getRow((short)11);
	    cell = row.createCell(7);
	    cell.setCellStyle(coloredtoprow);
		cell.setCellValue(" ");
		
		row = sheet.getRow((short)11);
	    cell = row.createCell(8);
	    cell.setCellStyle(coloredtoprightcorner);
		cell.setCellValue(" ");
		
		row=sheet.getRow(12);
	    CellRangeAddress cellRangeAddress4 = new CellRangeAddress(12, 12, 2, 4);
	    sheet.addMergedRegion(cellRangeAddress4);
	    cell = row.createCell(2);
	    cell.setCellStyle(toprowmerged);
		cell.setCellValue("Unnamed Cash");
		
		row = sheet.getRow((short)12);
	    cell = row.createCell(3);
	    cell.setCellStyle(coloredtoprow);
		cell.setCellValue(" ");
		
		row = sheet.getRow((short)12);
	    cell = row.createCell(4);
	    cell.setCellStyle(coloredtoprightcorner);
		cell.setCellValue(" ");
		
		row=sheet.getRow(12);
	    CellRangeAddress cellRangeAddress5 = new CellRangeAddress(12, 12, 6, 8);
	    sheet.addMergedRegion(cellRangeAddress5);
	    cell = row.createCell(6);
	    cell.setCellStyle(toprowmerged);
		cell.setCellValue("Named Cash");
		
		row = sheet.getRow((short)12);
	    cell = row.createCell(7);
	    cell.setCellStyle(coloredtoprow);
		cell.setCellValue(" ");
		
		row = sheet.getRow((short)12);
	    cell = row.createCell(8);
	    cell.setCellStyle(coloredtoprightcorner);
		cell.setCellValue(" ");
		
		sheet.getRow((short)13).createCell(2).setCellValue(createHelper.createRichTextString("$0.01"));
		sheet.getRow((short)13).createCell(3).setCellValue(jt1c.getText());
		row = sheet.getRow((short)13);
		cell = row.createCell(4);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("0.01*D14");			
		
		sheet.getRow((short)14).createCell(2).setCellValue(createHelper.createRichTextString("$0.05"));
		sheet.getRow((short)14).createCell(3).setCellValue(jt5c.getText());
		row = sheet.getRow((short)14);
		cell = row.createCell(4);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("0.05*D15");		
		
		sheet.getRow((short)15).createCell(2).setCellValue(createHelper.createRichTextString("$0.10"));
		sheet.getRow((short)15).createCell(3).setCellValue(jt10c.getText());
		row = sheet.getRow((short)15);
		cell = row.createCell(4);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("0.10*D16");		
		
		sheet.getRow((short)16).createCell(2).setCellValue(createHelper.createRichTextString("$0.25"));
		sheet.getRow((short)16).createCell(3).setCellValue(jt25c.getText());
		row = sheet.getRow((short)16);
		cell = row.createCell(4);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("0.25*D17");
		
		sheet.getRow((short)17).createCell(2).setCellValue(createHelper.createRichTextString("$0.50"));
		sheet.getRow((short)17).createCell(3).setCellValue(jt50c.getText());
		row = sheet.getRow((short)17);
		cell = row.createCell(4);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("0.50*D18");			
		
		sheet.getRow((short)18).createCell(2).setCellValue(createHelper.createRichTextString("$1.00"));
		sheet.getRow((short)18).createCell(3).setCellValue(jt1.getText());
		row = sheet.getRow((short)18);
		cell = row.createCell(4);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("1.00*D19");			
		
		sheet.getRow((short)19).createCell(2).setCellValue(createHelper.createRichTextString("$2.00"));
		sheet.getRow((short)19).createCell(3).setCellValue(jt2.getText());
		row = sheet.getRow((short)19);
		cell = row.createCell(4);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("2.00*D20");		
		
		sheet.getRow((short)20).createCell(2).setCellValue(createHelper.createRichTextString("$5.00"));
		sheet.getRow((short)20).createCell(3).setCellValue(jt5.getText());
		row = sheet.getRow((short)20);
		cell = row.createCell(4);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("5.0*D21");			
		
		sheet.getRow((short)21).createCell(2).setCellValue(createHelper.createRichTextString("$10.00"));
		sheet.getRow((short)21).createCell(3).setCellValue(jt10.getText());
		row = sheet.getRow((short)21);
		cell = row.createCell(4);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("10.0*D22");		
		
		sheet.getRow((short)22).createCell(2).setCellValue(createHelper.createRichTextString("$20.00"));
		sheet.getRow((short)22).createCell(3).setCellValue(jt20.getText());
		row = sheet.getRow((short)22);
		cell = row.createCell(4);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("20.0*D23");		
		
		sheet.getRow((short)23).createCell(2).setCellValue(createHelper.createRichTextString("$50.00"));
		sheet.getRow((short)23).createCell(3).setCellValue( jt50.getText());
		row = sheet.getRow((short)23);
		cell = row.createCell(4);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("50.0*D24");		
		
		sheet.getRow((short)24).createCell(2).setCellValue(createHelper.createRichTextString("$100.00"));
		sheet.getRow((short)24).createCell(3).setCellValue(jt100.getText());
		row = sheet.getRow((short)24);
		cell = row.createCell(4);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("100.0*D25");
	
		sheet.getRow((short)25).createCell(2).setCellValue(createHelper.createRichTextString("Unnamed cash total"));
		row = sheet.getRow((short)25);
		cell = row.createCell(4);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUM(E14:E25)");
		
		// named cash column 
		sheet.getRow((short)13).createCell(6).setCellValue(createHelper.createRichTextString("$0.01"));
		row = sheet.getRow((short)13);
		cell = row.createCell(8);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("0.01*H14");			

		sheet.getRow((short)14).createCell(6).setCellValue(createHelper.createRichTextString("$0.05"));
		row = sheet.getRow((short)14);
		cell = row.createCell(8);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("0.05*H15");		
		
		sheet.getRow((short)15).createCell(6).setCellValue(createHelper.createRichTextString("$0.10"));
		row = sheet.getRow((short)15);
		cell = row.createCell(8);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("0.10*H16");		

		sheet.getRow((short)16).createCell(6).setCellValue(createHelper.createRichTextString("$0.25"));
		row = sheet.getRow((short)16);
		cell = row.createCell(8);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("0.25*H17");	

		sheet.getRow((short)17).createCell(6).setCellValue(createHelper.createRichTextString("$0.50"));
		row = sheet.getRow((short)17);
		cell = row.createCell(8);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("0.50*H18");			
		
		sheet.getRow((short)18).createCell(6).setCellValue(createHelper.createRichTextString("$1.00"));
		row = sheet.getRow((short)18);
		cell = row.createCell(8);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("1.00*H19");	

		sheet.getRow((short)19).createCell(6).setCellValue(createHelper.createRichTextString("$2.00"));
		row = sheet.getRow((short)19);
		cell = row.createCell(8);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("2.00*H20");	

		sheet.getRow((short)20).createCell(6).setCellValue(createHelper.createRichTextString("$5.00"));
		row = sheet.getRow((short)20);
		cell = row.createCell(8);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("5.00*H21");		

		sheet.getRow((short)21).createCell(6).setCellValue(createHelper.createRichTextString("$10.00"));
		row = sheet.getRow((short)21);
		cell = row.createCell(8);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("10.00*H22");
		
		sheet.getRow((short)22).createCell(6).setCellValue(createHelper.createRichTextString("$20.00"));
		row = sheet.getRow((short)22);
		cell = row.createCell(8);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("20.00*H23");

		sheet.getRow((short)23).createCell(6).setCellValue(createHelper.createRichTextString("$50.00"));
		row = sheet.getRow((short)23);
		cell = row.createCell(8);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("50.00*H24");
		
		sheet.getRow((short)24).createCell(6).setCellValue(createHelper.createRichTextString("$100.00"));
		row = sheet.getRow((short)24);
		cell = row.createCell(8);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("100.00*H25");
		
		sheet.getRow((short)25).createCell(6).setCellValue(createHelper.createRichTextString("Named cash total"));
		row = sheet.getRow((short)25);
		cell = row.createCell(8);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("SUM(I14:I25)");
		
		
		row = sheet.getRow((short)28);
	    cell = row.createCell(2);
	    cell.setCellStyle(coloredstyle);
		cell.setCellValue("Named Cash");
		
		row = sheet.getRow((short)29);
		cell = row.createCell(2);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("I26");

		row = sheet.getRow((short)28);
	    cell = row.createCell(3);
	    cell.setCellStyle(coloredstyle);
		cell.setCellValue("Unnamed Cash");
		
		row = sheet.getRow((short)29);
		cell = row.createCell(3);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("E26");
		
		row = sheet.getRow((short)28);
	    cell = row.createCell(4);
	    cell.setCellStyle(coloredstyle);
		cell.setCellValue("Cash Total");
		
		row = sheet.getRow((short)29);
		cell = row.createCell(4);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("C30+D30");

		row = sheet.getRow((short)28);
	    cell = row.createCell(5);
	    cell.setCellStyle(coloredstyle);
		cell.setCellValue("Difference");
		
		row = sheet.getRow((short)29);
		cell = row.createCell(5);
		cell.setCellStyle(uncoloredstyle);
		cell.setCellFormula("E30-E9");
		
    }
    
    public void createTreasurersReport( HSSFWorkbook wb, CreationHelper createHelper, Sheet treasurersReport) {
    	
    		// setup column widths
		treasurersReport.setColumnWidth(0, 3400); // category
		treasurersReport.setColumnWidth(1, 3400); // eft
		treasurersReport.setColumnWidth(2, 3400); // bank deposit
		treasurersReport.setColumnWidth(3, 3400); //  total
    	
	    // setup background colors for sheet
	    HSSFPalette palette = wb.getCustomPalette();
	    HSSFColor myColor = palette.findSimilarColor(255, 202, 146);
	    short palIndex = myColor.getIndex();
	    
	    // setup the cell style for currency entries
	    CellStyle styleData = wb.createCellStyle();
	    styleData.setAlignment(HorizontalAlignment.RIGHT);
	    styleData.setDataFormat(wb.createDataFormat().getFormat(BuiltinFormats.getBuiltinFormat(8)));
    	
	    CellStyle coloredtoprow=wb.createCellStyle();
	    coloredtoprow.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    coloredtoprow.setFillForegroundColor(palIndex);
	    coloredtoprow.setBorderTop(BorderStyle.THIN);
	    coloredtoprow.setAlignment(HorizontalAlignment.CENTER);
    	
    		// Title row
    		Row row = treasurersReport.createRow((short)0);
	    row.setHeightInPoints(20);

	    Cell cell = row.createCell(0);
	    cell.setCellStyle(coloredtoprow);
	    cell.setCellValue(createHelper.createRichTextString("Category"));
	    
	    cell = row.createCell(1);
	    cell.setCellStyle(coloredtoprow);
	    cell.setCellValue(createHelper.createRichTextString("EFT"));
	    
	    cell = row.createCell(2);
	    cell.setCellStyle(coloredtoprow);
	    cell.setCellValue(createHelper.createRichTextString("Bank Deposit"));
	    
	    cell = row.createCell(3);
	    cell.setCellStyle(coloredtoprow);
	    cell.setCellValue(createHelper.createRichTextString("Total"));
	    
	    // Envelope row
		row = treasurersReport.createRow((short)1);
		row.setHeightInPoints(20);

		cell = row.createCell(0);
		cell.setCellValue(createHelper.createRichTextString("Envelope"));
    
		cell = row.createCell(1);
		cell.setCellStyle(styleData);
		cell.setCellFormula("Totals!G8");
    
		cell = row.createCell(2);
		cell.setCellStyle(styleData);
		cell.setCellFormula("Totals!F8");
		
		cell = row.createCell(3);
		cell.setCellStyle(styleData);
		cell.setCellFormula("Totals!H8");
	    
	    // Plate row
		row = treasurersReport.createRow((short)2);
		row.setHeightInPoints(20);

		cell = row.createCell(0);
		cell.setCellValue(createHelper.createRichTextString("Plate"));
    
		cell = row.createCell(1);
		cell.setCellStyle(styleData);
		cell.setCellFormula("Totals!G6");
    
		cell = row.createCell(2);
		cell.setCellStyle(styleData);
		cell.setCellFormula("Totals!F6");
		
		cell = row.createCell(3);
		cell.setCellStyle(styleData);
		cell.setCellFormula("Totals!H6");
	    
	    // Designated Ministries
		int i=1;
		int count = form.getDescriptionField().getItemCount();
		if (count >1) {
			while( i< count) {
				row = treasurersReport.createRow((short)2+i);
				row.setHeightInPoints(20);
				String s = form.getDescriptionField().getItemAt(i);
				cell = row.createCell(0);
				cell.setCellValue(createHelper.createRichTextString(s));
				
				cell = row.createCell(1);
				cell.setCellStyle(styleData);
				cell.setCellFormula("SUMIF(Data!I:I,\""+s+"\",Data!G:G)");
		    
				cell = row.createCell(2);
				cell.setCellStyle(styleData);
				cell.setCellFormula("SUMIF(Data!I:I,\""+s+"\",Data!E:E)+SUMIF(Data!I:I,\""+s+"\",Data!F:F)");
						
				cell = row.createCell(3);
				cell.setCellStyle(styleData);
				cell.setCellFormula("SUMIF(Data!I:I,\""+s+"\",Data!E:E)+SUMIF(Data!I:I,\""+s+"\",Data!F:F)+SUMIF(Data!I:I,\""+s+"\",Data!G:G)");
				
				i++;
			}
		}
		System.out.println("i="+i);
		row = treasurersReport.createRow((short)2+i);
		row.setHeightInPoints(20);
		cell = row.createCell(0);
		//cell.setCellStyle(styleData);
		cell.setCellValue(createHelper.createRichTextString("Total"));
		
		cell = row.createCell(1);
		cell.setCellStyle(styleData);
		cell.setCellFormula("SUM(B2:B"+(i+2)+")");
		
		cell = row.createCell(2);
		cell.setCellStyle(styleData);
		cell.setCellFormula("SUM(C2:C"+(i+2)+")");
		
		cell = row.createCell(3);
		cell.setCellStyle(styleData);
		cell.setCellFormula("SUM(D2:D"+(i+2)+")");
	    
    }
    
    public void createCheckReport(HSSFWorkbook wb, CreationHelper createHelper, Sheet checkReport) {
    	
    	// setup column widths
    			checkReport.setColumnWidth(0, 3400); // 
    			checkReport.setColumnWidth(1, 3400); // 
    			checkReport.setColumnWidth(2, 3400); // 
    			checkReport.setColumnWidth(3, 3400); //  
    	    	
    		    // setup background colors for sheet
    		    HSSFPalette palette = wb.getCustomPalette();
    		    HSSFColor myColor = palette.findSimilarColor(255, 202, 146);
    		    short palIndex = myColor.getIndex();
    		    
    		    // setup the cell style for currency entries
    		    CellStyle styleData = wb.createCellStyle();
    		    styleData.setAlignment(HorizontalAlignment.RIGHT);
    		    styleData.setDataFormat(wb.createDataFormat().getFormat(BuiltinFormats.getBuiltinFormat(8)));
    	    	
    		    CellStyle coloredtoprow=wb.createCellStyle();
    		    coloredtoprow.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    		    coloredtoprow.setFillForegroundColor(palIndex);
    		    coloredtoprow.setBorderTop(BorderStyle.THIN);
    		    coloredtoprow.setAlignment(HorizontalAlignment.CENTER);
    	    	
    	    		// Title row
    	    		Row row = checkReport.createRow((short)0);
    		    row.setHeightInPoints(20);

    		    Cell cell = row.createCell(0);
    		    cell.setCellStyle(coloredtoprow);
    		    cell.setCellValue(createHelper.createRichTextString("Check Number"));
    		    
    		    cell = row.createCell(1);
    		    cell.setCellStyle(coloredtoprow);
    		    cell.setCellValue(createHelper.createRichTextString("Check Amount"));
    		    
    		    int rowNumber=1;
    		    for(Donation d: offering) {
    		    		if(d.getCategory().equals("Check")) {
    		    			row = checkReport.createRow((short)rowNumber);
    		    			
    		    			System.out.println("row is " + rowNumber);
    		    			cell = row.createCell(0);
    		    			cell.setCellValue(rowNumber);
    		    			
    		    			cell = row.createCell(1);
    		    			cell.setCellStyle(styleData);
    		    			cell.setCellValue(d.getAmount());
    		    			
    		    			rowNumber++;
    		    			
    		    		}
    		    }
    	    
    	
    }
    

}
