import java.util.ArrayList;

import javax.swing.SwingUtilities;

public class ChurchOfferingRunner {

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		SwingUtilities.invokeLater(new Runnable() {
			public void run() {
				
				try {
					createAndShowGUI();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}
	
	private static void biggest(ArrayList a) {
		
	}
	
	private static void createAndShowGUI() throws Exception{
		new DataEntryForm();
	}
	
}
