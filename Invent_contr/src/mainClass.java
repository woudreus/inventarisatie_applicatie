
import javax.swing.UIManager;
import javax.swing.SwingUtilities;

	
	
	 

public class mainClass{
	
public static void main(String args[]) {
	 try {
	        UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
	    } catch (Exception ex) {
	        ex.printStackTrace();
	    }
	           
	           SwingUtilities.invokeLater(new Runnable() {
	               @Override
	               public void run() {
	                   new Gui().setVisible(true);
	               }
	   });     
	           
	           
}
	           
	    
   }


