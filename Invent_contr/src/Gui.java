import javax.swing.*;
import java.io.File;
import java.io.FileOutputStream;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Collections;
import java.util.Vector;
import java.awt.event.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigInteger;
import java.util.HashSet;
import java.awt.*;


public  class Gui extends JFrame {
  
	  /**
	  VERSION: ALPHA 01.00 
	 **/
	
	private static final long serialVersionUID = 1L;

	  JLabel select_inv = new JLabel("Selecteer Inventarisatie Controle:");
	  
	  JLabel label_rap = new JLabel("Rapport gemaakt door:");
	  JTextField textField_rap = new JTextField(30);

	  JLabel label_bewerker = new JLabel("Vergunning houder:");
	  JTextField textField_bewerker = new JTextField(30);
	  


      
	  JLabel label_gps = new JLabel("GPS ingeleverd door:");
	  JTextField textField_gps = new JTextField(30);

     
	  JLabel label_dat = new JLabel("Datum indiening GPS:");
	  JTextField textField_dat = new JTextField(30);
  
	  
      JLabel label_reg = new JLabel("Regio:");
	  JTextField textField_reg = new JTextField(30);
   
	  
      JLabel label_co = new JLabel("Coordinator: ");
	  JTextField textField_co = new JTextField(30);
	  
	  JLabel selectKV = new JLabel("Kapvak: ");
	  JLabel selectTer = new JLabel("Terrein: ");
     
	  
	  JLabel selectKV1 = new JLabel("Kapvak: ");
	  JLabel selectTer1 = new JLabel("Terrein: ");
     
	  JLabel dateMain = new JLabel("Datum 1: ");
	  JLabel dateMain2 = new JLabel("Datum 2: ");
	  
	  JLabel dateYear = new JLabel("YYYY ");
	  JLabel dateMonth = new JLabel("MM ");
	  JLabel dateDay = new JLabel("DD ");
	  TextField dateTYear = new TextField(5);
	  
	  JLabel labelInventLijn = new JLabel("Inventarisatie Lijn in meters: ");
	  TextField inventLijn = new TextField(10);
	  
	  Label opp = new Label();
	  
	  Label kvOpp[] = new Label[30];
	  TextField kvOppTF[] = new TextField[30];
	 
	  Vector<String> datainventdati = new Vector<String>();
	  
	  JLabel yearGPSLabel = new JLabel("YYYY ");
	  JLabel monthGPSLabel = new JLabel("MM ");
	  JLabel dayGPSlabel = new JLabel("DD ");
	  TextField yearGPS = new TextField(3);
	  TextField monthGPS = new TextField(3);
	  TextField dayGPS = new TextField(3);
	  
	  TextField dateTMonth = new TextField(5);
	  TextField dateTDay = new TextField(5);
	  
	  JLabel dateYear2 = new JLabel("YYYY ");
	  JLabel dateMonth2 = new JLabel("MM ");
	  JLabel dateDay2 = new JLabel("DD ");
	  TextField dateTYear2 = new TextField(5);
	  TextField dateTMonth2 = new TextField(5);
	  TextField dateTDay2 = new TextField(5);
	  
	  TextField Terrein = new TextField(8);
	
	  JLabel Terreinlabel = new JLabel("Terrein: ");
	  
	  JButton filterButton  = new JButton ("Filter");
	  JButton refreshButton = new JButton ("Refresh");
	  JButton CreateButton  = new JButton ("Vervaardig Rapport");
	
	  JButton continueButton = new JButton("Verder");
	  
	  JButton browse = new JButton("Zoeken");
	  JButton save = new JButton("Save As");
	  

	  JCheckBox sleepweg[] = new JCheckBox[30];
	  JCheckBox GIS[] = new JCheckBox[30];
	  JCheckBox hisProd[] = new JCheckBox[30];
	  
	  Label yearSubLabel[] = new Label[330];
	  Label monthSubLabel[] = new Label[30];
	  Label daySublabel[] = new Label[30];
	  TextField yearSub[] = new TextField[30];
	  TextField monthSub[] = new TextField[30];
	  TextField daySub[] = new TextField[30];
	  
	  Vector<String> indieningsDatum = new Vector<String>();
	  
	  JLabel labelImageSel = new JLabel("Overzichtskaart: ");
	  TextField imageSel = new TextField(72);
	  
	  JLabel saveLabel = new JLabel("Bestand opslaan als: ");
	  TextField SaveStringText	 = new TextField(72);
	  
	  JLabel sTerrein  = new JLabel(); 
	  JLabel sTerrein1 = new JLabel(); 
	
	 JLabel kapvakCB[] = new JLabel[30];
	
	  JList listboxc = new JList();	
	  JList listboxi = new JList();
		  
	  String fullList = new String();
	  String fullList2 = new String();	  
	  String terrJDBC = new String();
	  Double oppervlakte = null;
	  
	  Image image;
	  JFrame f = new JFrame();
	  JPanel p = new JPanel(); 
	  
	  JLabel stringDisplay0 = new JLabel();
	  Integer completion = new Integer(0);
	  
      public Gui() {
    	  
    	  GridBagConstraints cp = new GridBagConstraints();
		  cp.anchor = GridBagConstraints.FIRST_LINE_START;
		  cp.insets = new Insets(2, 2, 10, 2);
		  cp.fill = GridBagConstraints.FIRST_LINE_START;
		  cp.gridx = 0;
	      cp.gridy = 0;
    
         
          f.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
          f.setUndecorated(true);
          
          f.getContentPane().add(p);
          ImageIcon icon = new ImageIcon((getClass().getClassLoader().getResource("pleasewait_invent.jpg")));
		  JLabel label = new JLabel(icon);
		  
		
		  p.add(label, cp);
		 
		   f.setLocationRelativeTo(null);
		   
		   f.setVisible(false);
		   f.toFront();
		   f.requestFocus();
		   f.repaint();
		   toBack();
		   f.setSize(500, 333);
		   f.pack();
		  
		   completion = 10;
		   
		   if(completion == 10){
		   
		   stringDisplay0.setText("starting engines.");
		   p.add(stringDisplay0, cp);
    	  
		   }
    	listboxc.setSelectionMode(DefaultListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
		listboxc.setVisibleRowCount(4);
		JScrollPane panec = new JScrollPane(listboxc);
		panec.setPreferredSize(new Dimension(100, 80));
		  
		listboxi.setSelectionMode(DefaultListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
		listboxi.setVisibleRowCount(4);
		JScrollPane panei = new JScrollPane(listboxi);
		panei.setPreferredSize(new Dimension(100, 80));
	  	
	    JPanel parent = new JPanel();
	    parent.setLayout(new GridBagLayout());
	    add(parent);
	    

	    
	    GridBagConstraints constraints0 = new GridBagConstraints();
        constraints0.anchor = GridBagConstraints.FIRST_LINE_START;
        constraints0.insets = new Insets(2, 2, 2, 2);
	  	
        JPanel layout = new JPanel(new GridBagLayout());
        
        JPanel layout1 = new JPanel(new GridBagLayout());
        
        JPanel layout2 = new JPanel(new GridBagLayout());
        
        JPanel layout3 = new JPanel(new GridBagLayout());
        
        JPanel layout4 = new JPanel(new GridBagLayout());
        
        JPanel layoutButton = new JPanel(new GridBagLayout());
        
        JPanel layoutPlanning = new JPanel(new GridBagLayout());
        
        JPanel layoutPlanObj[] = new JPanel[30]; 
        
        JPanel layout5 = new JPanel(new GridBagLayout());
        
        JPanel layoutImageSel = new JPanel(new GridBagLayout());
        

 
        
        GridBagConstraints constraints = new GridBagConstraints();
        constraints.anchor = GridBagConstraints.FIRST_LINE_START;
        constraints.insets = new Insets(2, 2, 15, 2);
        
        constraints.gridx = 0;
        constraints.gridy = 0;     
        layout.add(dateMain, constraints);
        
        constraints.gridx = 0;
        constraints.gridy = 1;     
        layout.add(dateMain2, constraints);
        
        constraints.gridx = 1;
        constraints.gridy = 0;     
        layout.add(dateDay, constraints);
        
        constraints.gridx = 1;
        constraints.gridy = 1;
        layout.add(dateDay2, constraints);
        
        constraints.gridx = 2;
        constraints.gridy = 1;
        layout.add(dateTDay2, constraints);

        constraints.gridx = 3;
        constraints.gridy = 1;
        layout.add(dateMonth2, constraints);

        constraints.gridx = 4;
        constraints.gridy = 1;
        layout.add(dateTMonth2, constraints);
        
        constraints.gridx = 5;
        constraints.gridy = 1;
        layout.add(dateYear2, constraints);
        
        constraints.gridx = 6;
        constraints.gridy = 1;
        layout.add(dateTYear2, constraints);
        
        constraints.gridx = 2;
        constraints.gridy = 0;     
        layout.add(dateTDay, constraints);
        
        constraints.gridx = 3;
        constraints.gridy = 0;  
        layout.add(dateMonth, constraints);
        
        constraints.gridx = 4; 
        constraints.gridy = 0; 
        layout.add(dateTMonth, constraints);
        
        constraints.gridx = 5;
        constraints.gridy = 0; 
        layout.add(dateYear, constraints);
        
        constraints.gridx = 6;
        constraints.gridy = 0;
        layout.add(dateTYear, constraints);
        
        constraints.gridx = 0;
        constraints.gridy = 2;    
        layout.add(Terreinlabel, constraints);
        
        constraints.gridx = 1;
        constraints.gridy = 2;  
        constraints.gridwidth = 2;
        layout.add(Terrein, constraints);
        
        constraints.gridx = 5;
        constraints.gridy = 2;
        constraints.gridwidth = 2;
        constraints.insets = new Insets(0,-30,10,0);
        constraints.ipadx = 30;
        constraints.ipady = 11;
        layout.add(filterButton, constraints);
            
        layout.setBorder(BorderFactory.createTitledBorder(
               BorderFactory.createEtchedBorder(), "Filter Informatie"));    
        
        constraints0.gridx = 0;
        constraints0.gridy = 0;

        parent.add(layout, constraints0);
        
        GridBagConstraints constraints1 = new GridBagConstraints();
        constraints1.anchor = GridBagConstraints.FIRST_LINE_START;
        constraints1.insets = new Insets(2, 2, 10, 2);
        constraints1.fill = GridBagConstraints.FIRST_LINE_START;
        
       
        constraints1.gridx = 0;
        constraints1.gridy = 1;
        layout1.add(selectKV, constraints1);
        
        constraints1.gridx = 0;
        constraints1.gridy = 0;
        layout1.add(selectTer, constraints1); 
        
        constraints1.gridx = 1;
        constraints1.gridy = 0;
        layout1.add(sTerrein, constraints1); 
        
        constraints1.gridx = 1;
        constraints1.gridy = 1;
        layout1.add(panec, constraints1);
        
        layout1.setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createEtchedBorder(), "Filter Controle"));   
        
       constraints0.gridx = 1;
       constraints0.gridy = 0;
       parent.add(layout1, constraints0);
       
       GridBagConstraints constraints2 = new GridBagConstraints();
       constraints2.anchor = GridBagConstraints.FIRST_LINE_START;
       constraints2.insets = new Insets(2, 2, 10, 2);
       constraints2.fill = GridBagConstraints.FIRST_LINE_START;
       
       constraints2.gridx = 0;
       constraints2.gridy = 0;
       layout2.add(selectTer1, constraints2);
       
       constraints2.gridx = 1;
       constraints2.gridy = 0;
       layout2.add(sTerrein1, constraints2);
       
       constraints2.gridx = 0;
       constraints2.gridy = 1;
       layout2.add(selectKV1, constraints2);
       
       constraints2.gridx = 1;
       constraints2.gridy = 1;
       constraints2.ipadx = 20;
       layout2.add(panei, constraints2);
       
       layout2.setBorder(BorderFactory.createTitledBorder(
               BorderFactory.createEtchedBorder(), "Filter Inventarisatie"));   
       
      constraints0.gridx = 2;
      constraints0.gridy = 0;
      parent.add(layout2, constraints0);
        
      GridBagConstraints constraintsCButton = new GridBagConstraints();
      constraintsCButton.anchor = GridBagConstraints.CENTER;
      constraintsCButton.insets = new Insets(2, 2, 13, 2);
      
      constraintsCButton.gridx = 0;
      constraintsCButton.gridy = 0;
      layoutButton.add(continueButton, constraintsCButton);
      
      constraints0.gridx = 1;
      constraints0.gridy= 1;
      parent.add(layoutButton, constraints0);  
      
      GridBagConstraints constraints3 = new GridBagConstraints();
      constraints3.anchor = GridBagConstraints.FIRST_LINE_START;
      constraints3.insets = new Insets(2, 2, 10, 2);
      constraints3.fill = GridBagConstraints.FIRST_LINE_START;
      
      constraints3.gridx = 0;
      constraints3.gridy = 0;
      layout3.add(label_rap, constraints3);
      
      constraints3.gridx = 1;
      constraints3.gridy = 0;
      constraints3.ipadx= 40;
      constraints3.gridwidth = 3;
      layout3.add(textField_rap, constraints3);
      
      constraints3.gridx = 0;
      constraints3.gridy = 1;
      layout3.add(label_gps, constraints3);   
      
      constraints3.gridx = 1;
      constraints3.gridy = 1;
      layout3.add(textField_gps, constraints3);     
      
      constraints3.gridx = 0;
      constraints3.gridy = 2;
      layout3.add(label_dat, constraints3);      
      
      constraints3.gridx = 1;
      constraints3.gridy = 2;
      layout3.add(textField_dat, constraints3);      
      
      constraints3.gridx = 0;
      constraints3.gridy = 3;
      layout3.add(label_bewerker, constraints3);      
      
      constraints3.gridx = 1;
      constraints3.gridy = 3;
      layout3.add(textField_bewerker, constraints3);
      
      constraints3.gridx = 0;
      constraints3.gridy = 4;
      layout3.add(label_reg, constraints3);      
      
      constraints3.gridx = 1;
      constraints3.gridy = 4;
      layout3.add(textField_reg, constraints3);
            
      constraints3.gridx = 0;
      constraints3.gridy = 5;
      layout3.add(label_co, constraints3);
      
      constraints3.gridx = 1;
      constraints3.gridy = 5;
      layout3.add(textField_co, constraints3);      
      
      constraints0.gridx = 0;
      constraints0.gridy = 1;
      constraints0.gridwidth = 4;
      constraints0.ipadx= 10;
      parent.add(layout3, constraints0);
      
   
      layout3.setBorder(BorderFactory.createTitledBorder(
              BorderFactory.createEtchedBorder(), "Algemene Informatie"));   
      
      GridBagConstraints constraints4 = new GridBagConstraints();
      constraints4.anchor = GridBagConstraints.FIRST_LINE_START;
      constraints4.insets = new Insets(2, 2, 10, 2);
      constraints4.fill = GridBagConstraints.FIRST_LINE_START;
      
      constraints4.gridx = 0;
      constraints4.gridy = 0;
      layout4.add(refreshButton, constraints4);

      constraints4.gridx = 1;
      constraints4.gridy = 0;
      layout4.add(CreateButton, constraints4);
      
      constraints0.gridx = 0;
      constraints0.gridy = 4;
      parent.add(layout4, constraints0);

      GridBagConstraints constraints5 = new GridBagConstraints();
      constraints5.anchor = GridBagConstraints.FIRST_LINE_START;
      constraints5.insets = new Insets(2, 10, 10, 2);
      constraints5.fill = GridBagConstraints.FIRST_LINE_START;
      
      constraints5.gridx=0;
      constraints5.gridy=0;
      layout5.add(labelInventLijn, constraints5);
      
      constraints5.gridx=0;
      constraints5.gridy=1;
      layout5.add(inventLijn, constraints5);
      
      GridBagConstraints constraintsImageSel = new GridBagConstraints();
      constraintsImageSel.anchor = GridBagConstraints.FIRST_LINE_START;
      constraintsImageSel.insets = new Insets(2, 2, 10, 2);
      constraintsImageSel.fill = GridBagConstraints.FIRST_LINE_START;
      
      constraintsImageSel.gridx = 0;
      constraintsImageSel.gridy = 0;
      layoutImageSel.add(labelImageSel, constraintsImageSel);
      
      constraintsImageSel.gridx = 1;
      constraintsImageSel.gridy = 0;
      layoutImageSel.add(imageSel, constraintsImageSel);

      
      constraintsImageSel.gridx = 2;
      constraintsImageSel.gridy = 0;
      layoutImageSel.add(browse, constraintsImageSel);
      
      constraintsImageSel.gridx = 0;
      constraintsImageSel.gridy = 1;
      layoutImageSel.add(saveLabel,  constraintsImageSel);
      
      constraintsImageSel.gridx = 1;
      constraintsImageSel.gridy = 1;
      layoutImageSel.add(SaveStringText, constraintsImageSel);
     	 
      
      constraintsImageSel.gridx = 2;
      constraintsImageSel.gridy = 1;
      layoutImageSel.add(save, constraintsImageSel);
            
      constraints0.gridx = 0;
      constraints0.gridy = 3;
      parent.add(layoutImageSel, constraints0);
      
      layoutImageSel.setBorder(BorderFactory.createTitledBorder(
              BorderFactory.createEtchedBorder(), "Select/Save"));  
      
     browse.addActionListener(new ActionListener() {
          public void actionPerformed(ActionEvent ae) {
            JFileChooser fileChooser = new JFileChooser();
            int returnValue = fileChooser.showOpenDialog(null);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
              File selectedFile = fileChooser.getSelectedFile();
             String imageString = selectedFile.getAbsolutePath() ;

            	 imageSel.setText(imageString);
            }
          }
        });
         
 
     save.addActionListener(new ActionListener(){
    	 public void actionPerformed(ActionEvent ae) {
     
    		 JFileChooser fileChooser = new JFileChooser();
             int returnValue = fileChooser.showSaveDialog(null);
    	      if (returnValue == JFileChooser.APPROVE_OPTION) {
                 File selectedFile = fileChooser.getSelectedFile();
                String SaveString = selectedFile.getAbsolutePath() ;
               	
                if(!SaveString.endsWith(".docx") ){
                	
                	String ReformatSaveString = String.format("%s%s", SaveString, ".docx");

                	SaveStringText.setText(ReformatSaveString);
                	
                } else if (SaveString.endsWith(".docx")){
                 SaveStringText.setText(SaveString);
                }    
    	      }
     }
    	 }
     );
     
 
      inventLijn.addTextListener(new CustomTextListener() {
    	  
    	  public void textValueChanged(TextEvent e) {  
    	  
    		try{  
    	  Double lengteLijn = Double.parseDouble(inventLijn.getText() );
    	  
    	 Double oppervlakte = (lengteLijn * 20)/10000;

    	  opp.setSize(150,20);
    	  opp.setText("Oppervlakte: " +  oppervlakte + "ha" );
    	  
    		} catch(Exception z){  }
    	  
    	  }   });
      
      constraints5.gridx=0;
      constraints5.gridy=3;
      layout5.add(opp, constraints5);
      
      constraints0.gridx=2;
      constraints0.gridy=1;
      parent.add(layout5, constraints0);
      
      Terrein.addTextListener(new CustomTextListener(){
    	
    	  public void textValueChanged(TextEvent e) { 
    		  
    		  try{
    			  
    			  sTerrein.setText(Terrein.getText());
    			  sTerrein1.setText(Terrein.getText());
    			  
    		  } catch(Exception zz){}
    		  
    	  }
    	  
      }
    		  );
      
      layout3.setVisible(false);
      layout4.setVisible(false);
      layout5.setVisible(false);
      layoutImageSel.setVisible(false);


        pack();
      	parent.setPreferredSize(new Dimension(680, 230));
      	parent.setMinimumSize(new Dimension(680, 230));
        setLocationRelativeTo(null);
  
      setResizable(false);
    
      setTitle("Inventarisatie Controle");
      setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        
      
      theFilter filter = new theFilter();
      filterButton.addActionListener(filter);
      
      theCreate create = new theCreate();
      CreateButton.addActionListener(create);
      
      window windowpane = new window();
      CreateButton.addActionListener(windowpane);  
      
      GridBagConstraints constraintsCB = new GridBagConstraints();
      constraintsCB.anchor = GridBagConstraints.FIRST_LINE_START;
      constraintsCB.insets = new Insets(5, 5, 5, 5);
      constraintsCB.fill = GridBagConstraints.FIRST_LINE_START;
 	  constraintsCB.ipadx = 1;
      
 	
 	GridBagConstraints constraintsPlan = new GridBagConstraints();
 	constraintsPlan.anchor = GridBagConstraints.FIRST_LINE_START;
 	constraintsPlan.insets = new Insets(2, 2, 2, 2);
 	constraintsPlan.fill = GridBagConstraints.FIRST_LINE_START;
 	constraintsPlan.ipadx = 1;
	

 	
      continueButton.addActionListener(new ActionListener() {
          @Override
          public void actionPerformed(ActionEvent e) {
        	  
        	  try{
        	  
        	  int numKvSelect = listboxi.getSelectedValuesList().size();
              if(numKvSelect < 30){
            	  
         layout3.setVisible(true);
         layout4.setVisible(true);
         layout5.setVisible(true);
         continueButton.setVisible(false);
         layoutImageSel.setVisible(true);
         layoutPlanning.setVisible(true);
    
        Object[] kvLabel = listboxi.getSelectedValuesList().toArray();
         

        for (int i= 0; i < numKvSelect; i++){       
                    
        	 layoutPlanObj[i] = new JPanel(new GridBagLayout()); 
             layoutPlanObj[i].setVisible(true);
                    
        	 sleepweg[i] = new JCheckBox("Sleepweg Gepland");
        	 constraintsCB.gridx = 0;
        	 constraintsCB.gridy = 0;
        	 constraintsCB.ipadx = 1;
        	 constraintsCB.gridwidth = 3;
        	 constraintsCB.insets = new Insets(2, 2, 2, 2);
        	 layoutPlanObj[i].add(sleepweg[i], constraintsCB);
        	 
        	 GIS[i] = new JCheckBox("GIS Data Ingediend");
        	 constraintsCB.gridx = 3;
        	 constraintsCB.gridy = 0;
        	 constraintsCB.gridwidth = 2;
        	 layoutPlanObj[i].add(GIS[i], constraintsCB);
        	 
        	 hisProd[i] = new JCheckBox("Historische Productie");
        	 constraintsCB.gridx = 5;
        	 constraintsCB.gridy = 0;
             layoutPlanObj[i].add(hisProd[i], constraintsCB); 
             
             yearSubLabel[i] = new Label("Ingediend d/m/y");
        	 yearSubLabel[i].setSize(10,10);
        	 constraintsCB.gridwidth = 1;
        	 constraintsCB.gridx = 0;
        	 constraintsCB.gridy = 1;
        	 constraintsCB.insets = new Insets(0, 2, 1, 2);
        	 constraintsCB.ipadx = 1;
        	 layoutPlanObj[i].add(yearSubLabel[i], constraintsCB);
        	 
        	 yearSub[i] = new TextField(4);
           	 constraintsCB.ipadx = 0;
        	 constraintsCB.gridx = 3;
        	 constraintsCB.gridy = 1;
        	 layoutPlanObj[i].add(yearSub[i], constraintsCB);
        	 
        	 monthSub[i] = new TextField(2);

        	 constraintsCB.gridx = 2;
        	 constraintsCB.gridy = 1;
        	 layoutPlanObj[i].add(monthSub[i], constraintsCB);
        	 
        	 daySub[i] = new TextField(2);
        	 constraintsCB.gridx = 1;
        	 constraintsCB.gridy = 1;
        	 layoutPlanObj[i].add(daySub[i], constraintsCB);
             
              kvOpp[i] = new Label("Oppervlakte kapvak:");
              constraintsCB.gridx = 5;
              constraintsCB.gridy = 1;
              layoutPlanObj[i].add(kvOpp[i], constraintsCB);
              
       	      kvOppTF[i] = new TextField(4);
         	  constraintsCB.gridx = 6;
         	  constraintsCB.gridy = 1;
         	  layoutPlanObj[i].add(kvOppTF[i], constraintsCB);
         	          
         
         layoutPlanObj[i].setBorder(BorderFactory.createTitledBorder(
                 BorderFactory.createEtchedBorder(), "Kapvak " + kvLabel[i]));   
         
         constraintsPlan.gridx = 0;
         constraintsPlan.gridy = i;
         layoutPlanning.add(layoutPlanObj[i], constraintsPlan);
         }
         
         
        layoutPlanning.setBorder(BorderFactory.createTitledBorder(
                 BorderFactory.createEtchedBorder(), "Algemene Informatie Kapvak(ken)" ));           
        
        constraints0.gridx = 0;
        constraints0.gridy = 2;
        parent.add(layoutPlanning, constraints0);
        
        int sizeX = 570 + (80*numKvSelect);
        
        if (sizeX <= 980){
        	
        	parent.setPreferredSize(new Dimension(680, sizeX));
        	parent.setMinimumSize(new Dimension(680, sizeX));
        	pack();} else{
        		
        	       
         	parent.setPreferredSize(new Dimension(680, 960));
        	parent.setMinimumSize(new Dimension(680, 960));
            setLocationRelativeTo(null);
            
       	       	
         JScrollPane scrollPane = new JScrollPane(layoutPlanning);
         pack(); 
         scrollPane.setPreferredSize(new Dimension(630, 420));
         parent.add(scrollPane, constraints0);
      
                }
        
        } else {JOptionPane.showMessageDialog(null, "Let op! Te veel kapvakken geselecteerd!");}
        
        	  }catch(Exception ContinueError){JOptionPane.showMessageDialog(null, "Error: 675-786 Fout opgetreden: " + ContinueError.getMessage() );}
        	  
       }
     });
      
      
      inventLijn.addTextListener(new CustomTextListener(){
    	  
          public void textValueChanged(TextEvent e) { 
        	
        	  try{
        		   Integer.parseInt(inventLijn.getText()); 
        		    
        		   inventLijn.setBackground(Color.white);
        		   
        		   if( inventLijn.getText().equals("") ){inventLijn.setBackground(Color.white);}
        		   
          	} catch (Exception NaN){ inventLijn.setBackground(Color.RED); }
          	}
          });
      
      dateTDay.addTextListener(new CustomTextListener(){
    	  
          public void textValueChanged(TextEvent e) {        	  
        	  
              try {
            	  
            	  
            	  if ( !dateTDay.getText().isEmpty() ){
              	  Integer dayInt = Integer.parseInt(dateTDay.getText());       	
                         	  
                if ( dateTDay.getText().length() > 2| dateTDay.getText().length() == 1 & dayInt > 3 | dayInt > 31){dateTDay.setBackground(Color.red);}             
                  else if (dateTDay.getText() == " " ) {dateTDay.setBackground(Color.white);} else  {dateTDay.setBackground(Color.white);} 
                
                
            	  } else {dateTDay.setBackground(Color.white);} 
            	  
              	  } catch(Exception k) {
              		  		  dateTDay.setBackground(Color.red);
              	  }              	  
          }      	  
      });
      
      dateTYear.addTextListener(new CustomTextListener(){
    	          	  
    	        public void textValueChanged(TextEvent e) {        	  
        	  
        try {
        	
        	 if ( !dateTYear.getText().isEmpty() ){
           	  Integer yearInt = Integer.parseInt(dateTYear.getText());       	
                      	  
             if ( dateTYear.getText().length() > 4 & yearInt > 2100 | yearInt < 1900 & yearInt > 202){dateTYear.setBackground(Color.red);}           
          else if (dateTYear.getText() == " " ) {dateTYear.setBackground(Color.white);} else  {dateTYear.setBackground(Color.white);}  
        	
        	} else {dateTYear.setBackground(Color.white);}  
        	 
        	  } catch(Exception k) {
        		  		  dateTYear.setBackground(Color.red);
        	  }}}		  
    		  );
      
          
          
      dateTMonth.addTextListener(new CustomTextListener(){
          public void textValueChanged(TextEvent e) {

        		  
        	  try {
        		  
        			if ( !dateTMonth.getText().isEmpty() ){
        		  Integer monthInt = Integer.parseInt(dateTMonth.getText());      
        	   	
                       	  
              if ( dateTMonth.getText().length() > 2 | dateTMonth.getText().length() == 1 & monthInt > 2 | monthInt > 12  ){dateTMonth.setBackground(Color.red);}             
              else if (dateTMonth.getText() == " " ) {dateTMonth.setBackground(Color.white);} else  {dateTMonth.setBackground(Color.white);} 
              
        			} else {dateTMonth.setBackground(Color.white);} 
              
            	  } catch(Exception k) {
            		  dateTMonth.setBackground(Color.red);
            	  }}});
      
      
      
      dateTDay2.addTextListener(new CustomTextListener(){
    	  
          public void textValueChanged(TextEvent e) {        	  
        	  
              try {
            	  
            	  if ( !dateTDay2.getText().isEmpty() ){
            	  
              	  Integer dayInt = Integer.parseInt(dateTDay2.getText());       	
                         	  
                if ( dateTDay2.getText().length() > 2 |dateTDay2.getText().length() == 1 & dayInt > 3 | dayInt > 31){dateTDay2.setBackground(Color.red);}             
                else if (dateTDay2.getText() == " " ) {dateTDay2.setBackground(Color.white);} else  {dateTDay2.setBackground(Color.white);} 
                
            	  } else  {dateTDay2.setBackground(Color.white);} 
            	  
              	  } catch(Exception k) {
              		  		  dateTDay2.setBackground(Color.red);
              	  }
              	  
          }  
    	  
      });
      
      
      dateTYear2.addTextListener(new CustomTextListener(){
    	          	  
    	        public void textValueChanged(TextEvent e) {        	  
        	  
        try {
        	 if ( !dateTYear2.getText().isEmpty() ){
        	  Integer yearInt = Integer.parseInt(dateTYear2.getText());       	
                   	  
          if ( dateTYear2.getText().length() > 4 & yearInt > 2100 | yearInt < 1900 & yearInt > 202){dateTYear2.setBackground(Color.red);}             
          else if (dateTYear2.getText() == " " ) {dateTYear2.setBackground(Color.white);} else  {dateTYear2.setBackground(Color.white);} 
          
        	 } else  {dateTYear2.setBackground(Color.white);} 
        	 
        	  } catch(Exception k) {
        		  		  dateTYear2.setBackground(Color.red);
        	  }
        	  
    	        }}		  
    		  );
      
          
          
      dateTMonth2.addTextListener(new CustomTextListener(){
          public void textValueChanged(TextEvent e) {
        	         	  
        	  try {
        		  
        		  if ( !dateTMonth2.getText().isEmpty() ){
        		  
            	  Integer monthInt = Integer.parseInt(dateTMonth2.getText());       	
                       	  
              if ( dateTMonth2.getText().length() > 2 | dateTMonth2.getText().length() == 1 & monthInt > 2 | monthInt > 12  ){dateTMonth2.setBackground(Color.red);}             
              else if (dateTMonth2.getText() == " " ) {dateTMonth2.setBackground(Color.white);} else  {dateTMonth2.setBackground(Color.white);} 
              
        		  } else  {dateTMonth2.setBackground(Color.white);} 
              
            	  } catch(Exception k) {
            		  dateTMonth2.setBackground(Color.red);
            	  } }});
        
      
      imageSel.addTextListener(new CustomTextListener(){
      	  
    public void textValueChanged(TextEvent e) {    
    	
    	if (!imageSel.getText().endsWith(".jpg")  ){
    		
    		imageSel.setBackground(Color.RED);
    		
    	} else if(imageSel.getText().endsWith(".jpg")){imageSel.setBackground(Color.white);
    } else if(imageSel.getText().equals("")){imageSel.setBackground(Color.white);}
    }});
      
      SaveStringText.addTextListener(new CustomTextListener(){
      	  
    	    public void textValueChanged(TextEvent e) {    
    	    	
    	    	if (!SaveStringText.getText().endsWith(".docx")){
    	    		
    	    		 SaveStringText.setBackground(Color.RED);
    	    		
    	    	} else if(SaveStringText.getText().endsWith(".docx")){SaveStringText.setBackground(Color.white);}
    	    	 else if(SaveStringText.getText().equals("")){SaveStringText.setBackground(Color.white);}
    	    }});
      
      refreshButton.addActionListener(
    		  new ActionListener(){
   @Override
    		  public void actionPerformed(ActionEvent h){
    		
    			   
    				dateTYear.setText("");
    				dateTYear.setBackground(Color.white);
    				dateTYear2.setText("");
    				dateTYear2.setBackground(Color.white);
    				dateTMonth.setText("");
    				dateTMonth.setBackground(Color.white);
    				dateTMonth2.setText("");
    				dateTMonth2.setBackground(Color.white);
    				dateTDay.setText("");
    				dateTDay.setBackground(Color.white);
    				dateTDay2.setText("");
    				dateTDay2.setBackground(Color.white);
    				inventLijn.setText("");
    				inventLijn.setBackground(Color.white);
    				textField_rap.setText("");
    				textField_bewerker.setText("");
    				textField_gps.setText("");
    				textField_dat.setText("");
    				textField_reg.setText("");
    				textField_co.setText("");
    				Terrein.setText(""); 
    				labelImageSel.setText("");   				  
    				opp.setText(""); 
    				SaveStringText.setText("");
    				imageSel.setText("");
    				
		  
    				listboxc.setVisible(false);
    				listboxi.setVisible(false);
    
    				listboxc.clearSelection();
    				listboxi.clearSelection();

    				continueButton.setVisible(true);
    				
    				  layoutPlanning.removeAll();
      				  layout3.setVisible(false);
    			      layout4.setVisible(false);
    			      layout5.setVisible(false);
    			      layoutImageSel.setVisible(false);
    			      layoutPlanning.setVisible(false);
    			      layoutPlanObj[ listboxi.getSelectedValuesList().size()].removeAll();
    			
    			    pack();
    			   	parent.setPreferredSize(new Dimension(680, 230));
    		    	parent.setMinimumSize(new Dimension(680, 230));
    		        setLocationRelativeTo(null);
    		     
    				   }
    		   }
    		  );

      
 
      } //end of GUI
      
      class CustomTextListener implements TextListener {
          public void textValueChanged(TextEvent e) {      	  
        	  

        	  
        	  }
          
       }
      
	   private class theFilter implements ActionListener{
		   public void actionPerformed(ActionEvent f){
			   
			   try{
			   
			      listboxc.setVisible(true);
		    	  listboxi.setVisible(true);
				  
			 final	Vector<Integer> datai = new Vector<Integer>();
			 final	Vector<Integer> datac = new Vector<Integer>();
			
		    	String  terreinTextfield = Terrein.getText();
		    	String 	yearTextfield = dateTYear.getText();
		    	String	monthTextfield = dateTMonth.getText();
		    	String 	dayTextfield = dateTDay.getText();
		    	
		    	String 	yearTextfield2 = dateTYear2.getText();
		    	String	monthTextfield2 = dateTMonth2.getText();
		    	String 	dayTextfield2 = dateTDay2.getText();
		    	
		    	String date_db = String.format("'%s-%s-%s'", yearTextfield,monthTextfield,dayTextfield);
		    	String date_db2 = String.format("'%s-%s-%s'", yearTextfield2,monthTextfield2,dayTextfield2);
		    			
		    	String terrein_db = String.format("'%s'", terreinTextfield);
		    			    
		    	Connection c = null;
			  	Statement stmti = null;
			  	Statement stmtc = null;
	
		    try {
			  	       Class.forName("org.postgresql.Driver");
			  	       c = DriverManager
			  	          .getConnection("jdbc:postgresql://192.168.10.7:5432/sr",
			  	          "stephan", "playfairs");
			  	       c.setAutoCommit(false);
			  	      
			  	       stmtc = c.createStatement();
		   if( dayTextfield2.isEmpty()  && monthTextfield2.isEmpty() && yearTextfield2.isEmpty()){
			  	       ResultSet rsc = stmtc.executeQuery( "SELECT DISTINCT kapvak::integer, terrein "
			  	       										+ "FROM opname.honderd_procent_algemeen WHERE start_opname =" 
			  	       										+ date_db + "AND terrein =" + terrein_db +";" );
		   
			  	     while ( rsc.next() ) {
			  	   	  	   
			  	    	 Integer kapvak  = rsc.getInt("kapvak");	  	    
			  	    	   			  	    	 
			  	   
			  	            
			  	          datac.addElement(kapvak);
			  	       
			  	   	  	    		  }
			  		rsc.close();
		   
		   } else if (!dayTextfield2.isEmpty()  && !monthTextfield2.isEmpty() && !yearTextfield2.isEmpty()){
			      ResultSet rsc = stmtc.executeQuery( "SELECT DISTINCT kapvak::integer, terrein "
			  	       									+ "FROM opname.honderd_procent_algemeen WHERE start_opname >=" 
			  	       									+ date_db + "::date " + "AND start_opname <="+ date_db2 +"::date AND terrein =" + terrein_db +";" );
			      
			      while ( rsc.next() ) {
		  	   	  	   
		  	    	  Integer kapvak  = rsc.getInt("kapvak");	  	    
		  	    	   			  	    	 
		  	    	  String terrein = rsc.getString("terrein");
		  	      	  	            	  	          
		  	          Gui.this.terrJDBC = terrein;
		  	            
		  	          datac.addElement(kapvak);
		  	       	 Collections.sort(datac);
		  	       
		  	   	  	    		  }
			  	rsc.close();
		   }
			  	    		  	      
			  	       
			  	       stmti = c.createStatement();
					  	  	         	    	  	   
		   			   ResultSet rsi = stmti.executeQuery( "SELECT DISTINCT kapvak::integer FROM inventarisatie.plot_info"
		   			   										+ "  WHERE plot_info.terrein ="+ terrein_db +";" );
			  	       
			  	       while ( rsi.next() ) {
			  	   	  	   
			  	    	  Integer kapvak  = rsi.getInt("kapvak");
		 
			  	       	   	datai.addElement(kapvak);
			  	       	   	Collections.sort(datai);
			  	       	   	
			  	    		  }
			  	   
					   listboxc.setListData(datac);	
					   listboxi.setListData(datai);
					  			  	   
			  	       stmti.close();
			  	       stmtc.close();
			  	       c.close();
			  	       rsi.close();
		  	       
			  		  }
			  		  
			  		  catch ( Exception g ) {
			  		         System.err.println( g.getClass().getName()+": "+ g.getMessage() );
			  		    
			  	}
			   }catch(Exception FilterError){JOptionPane.showMessageDialog(null, "Error: 977-1076 Fout opgetreden: " + FilterError.getMessage() );}
			   }
	   }
	   
	   private class window implements ActionListener{
		   public void actionPerformed(ActionEvent h){

			   
			   f.setVisible(true);
			   f.toFront();
			   f.requestFocus();
			   f.repaint();
					   
		   	}
		   }
	   
	   private class theCreate implements ActionListener{
		   public void actionPerformed(ActionEvent h){
			   
			  
			 
			     int numKvSelect = listboxi.getSelectedValuesList().size();
			     
			     
			     
			 
			      
			      for (int i= 0; i < numKvSelect; i++){       
			          
			     	 indieningsDatum.addElement( String.format("%s-%s-%s",daySub[i].getText() , monthSub[i].getText(), yearSub[i].getText()  ) );
			      
			          }
			      
		
		   
				String terreinTextfield = Terrein.getText();
				String terrein_db = String.format("'%s'", terreinTextfield);
				String terrein_dbOP = String.format("%s", terreinTextfield);
				
		    	String 	yearTextfield = dateTYear.getText();
		    	String	monthTextfield = dateTMonth.getText();
		    	String 	dayTextfield = dateTDay.getText();
		    	
		    	String 	yearTextfield2 = dateTYear2.getText();
		    	String	monthTextfield2 = dateTMonth2.getText();
		    	String 	dayTextfield2 = dateTDay2.getText();
		    	
		    	String date_db = String.format("'%s-%s-%s'", yearTextfield,monthTextfield,dayTextfield);
		    	String date_db2 = String.format("'%s-%s-%s'", yearTextfield2,monthTextfield2,dayTextfield2);
				
		    	final String kv_sl_cont = new String(listboxc.getSelectedValuesList().toString());
				final String kv_sl_invent = new String(listboxi.getSelectedValuesList().toString());
				String	kvQuery_contOP = kv_sl_invent.substring(1, kv_sl_invent.length()-1);

				   int response =   JOptionPane.showConfirmDialog(null, String.format("%s %s %s%s %s", "Rapport voor terrein", terrein_dbOP, "KV: ",
						   kvQuery_contOP, "wordt vervaardigt.\n Is dit juist?"), "Bevestig",
					        JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE);
				   
					    if (response == JOptionPane.NO_OPTION) {
					       	f.setVisible(false);
					    	return;
					    } else if (response == JOptionPane.YES_OPTION) {
					
					    } else if (response == JOptionPane.CLOSED_OPTION) {
					    	f.setVisible(false);
					    	return;
					    }
		    	
				 final	Vector<String> datatopoi = new Vector<String>();
				 final	Vector<String> databodemi = new Vector<String>();
				 final	Vector<String> databosi = new Vector<String>();
				 final	Vector<String> datakvi = new Vector<String>();
			
				 final	Vector<String> dataodk 	   = new Vector<String>();
				 final	Vector<String> datapLeider = new Vector<String>();
				 final	Vector<String> databKenner = new Vector<String>();
				 final	Vector<String> dataopnemer3 = new Vector<String>();
				 final	Vector<String> dataopnemer4 = new Vector<String>();
				 final	Vector<String> datatopoc = new Vector<String>();
				 final	Vector<String> databosc = new Vector<String>();
				 final	Vector<String> databodemc = new Vector<String>();
				 final	Vector<String> dataondergroei = new Vector<String>();
				 final	Vector<String> dataopmc = new Vector<String>();
				 final	Vector<String> dataehka = new Vector<String>();
				 final  Vector<String> datakapvn = new Vector<String>();
				 final	Vector<String> datamarkmethod = new Vector<String>();
				 final	Vector<String> datakvc = new Vector<String>();
				 final	Vector<String> datanbosact = new Vector<String>();
				 final	Vector<String> datamarkkleur = new Vector<String>();
				 final	Vector<String> dataAfbakening = new Vector<String>();

				 final	Vector<String> datamark_afst = new Vector<String>();
				 final	Vector<String> data_ta = new Vector<String>();
				 final 	Vector<String> data_date = new Vector<String>();
				 final  Vector<String> dataJaarehka = new Vector<String>();
				 final  Vector<String> dataSchaalehka = new Vector<String>();
				 final  Vector<String> databoommark = new Vector<String>();
				 final  Vector<String> datasleepmark = new Vector<String>();
				 
				 final Vector<String> dataSuc = new Vector<String>();
				 final Vector<String> datakvctc = new Vector<String>();
				 final Vector<String> datapercc = new Vector<String>();
				 final Vector<String> databoomc = new Vector<String>();
				 final Vector<String> datahscklfout = new Vector<String>();
				 final Vector<String> datadbhc = new Vector<String>();
				 final Vector<String> datastrhc = new Vector<String>();
				 final Vector<String> datahcc = new Vector<String>();
				 final Vector<String> datakwac = new Vector<String>();
				 final Vector<String> datamark = new Vector<String>();
				 final Vector<String> dataStronk = new Vector<String>();
				 final Vector<String> databoomna = new Vector<String>();
				 
				 final Vector<String> dataSui = new Vector<String>();
				 final Vector<String> datakvcti = new Vector<String>();
				 final Vector<String> dataperci = new Vector<String>();
				 final Vector<String> databoomi = new Vector<String>();
				 final Vector<String> datahscgoed = new Vector<String>();
				 final Vector<String> datadbhi = new Vector<String>();
				 final Vector<String> datastrhi = new Vector<String>();
				 final Vector<String> datahci = new Vector<String>();
				 final Vector<String> datakwai = new Vector<String>();
				 final Vector<String> datasel = new Vector<String>();
				 final Vector<String> dataUIDc = new Vector<String>();
				 final Vector<String> dataopnemeri = new Vector<String>();
				 final Vector<String> datakvbtype = new Vector<String>();
				 
				 final Vector<String> dataTotalTrees = new Vector<String>();
					 
				 final Vector<Double> cub1 = new Vector<Double>();
				 final Vector<Double> cub2 = new Vector<Double>();
				 final Vector<Double> cub3 = new Vector<Double>();
				 
				 final Vector<Integer> dataDeltadbh = new Vector<Integer>();
				 final Vector<Integer> dataDeltahc = new Vector<Integer>();
				 
				 final Vector<Integer> dataKv_verb = new Vector<Integer>(); 
				 final Vector<String>  dataHs_verb = new Vector<String>();
				 final Vector<Integer> dataNum_verb= new Vector<Integer>();
				 
				 final Vector<Integer> dataKv_onderma = new Vector<Integer>(); 
				 final Vector<String>  dataHs_onderma = new Vector<String>();
				 final Vector<Integer> dataNum_onderma= new Vector<Integer>();
				 
				 final Vector<Integer> markselKV = new Vector<Integer>();
				 final Vector<Integer> mark1sel1 = new Vector<Integer>();
				 final Vector<Integer> mark0sel1 = new Vector<Integer>();
				 final Vector<Integer> mark1sel0 = new Vector<Integer>();
				 final Vector<Integer> mark0sel0 = new Vector<Integer>();
				  
				 final	HashSet<String> uniqueDate = new HashSet<String>();
				 final  HashSet<String> uniquepLeiders = new HashSet<String>();
				 final  HashSet<String> uniqueBoomKenners = new HashSet<String>();
				 final  HashSet<String> uniqueOpn1 = new HashSet<String>();	 
				 final  HashSet<String> uniqueOpn2 = new HashSet<String>();
				 
				 final Vector<String> datamarkmethodString = new Vector<String>();
				 final Vector<String> datasleepmarkString = new Vector<String>();
				 final Vector<String> databoommarkString = new Vector<String>();
				 final Vector<String> datamarkJaarString = new Vector<String>();
				 final Vector<String> datamarkSchaalString = new Vector<String>();
	
				 final Vector<String> datakapvnString = new Vector<String>();
				 final Vector<String> dataAfbString = new Vector<String>();
				 
				 String ehkaString = new String();
				 String nbosactString= new String();
				 
				 Integer countboomcodes = new Integer(0);

				 			 
		    	Connection c = null;
			  	Statement stmti_alg = null;
			  	Statement stmtc_alg = null;
			  	Statement stmti_tree = null;
			  	Statement stmtc_tree = null;
			  	Statement stmtc_Terrain = null;
			  	Statement stmtci = null;
			  	Statement stmttemp = null;
			  	Statement temp_alg = null;
			  	Statement  stmtemp1 = null;
			  	Statement stmttot = null;
			  	
			    GridBagConstraints cp = new GridBagConstraints();
				  cp.anchor = GridBagConstraints.FIRST_LINE_START;
				  cp.insets = new Insets(2, 2, 10, 2);
				  cp.fill = GridBagConstraints.FIRST_LINE_START;
				  cp.gridx = 0;
			      cp.gridy = 0;
			  	
			  	 completion = 20;
				   
				   if(completion == 20){
					   stringDisplay0.setText("adding sugar, spice, and everything nice");
					   p.add(stringDisplay0, cp);}
	
					
					   
			
				
			  	
			 FileOutputStream out = null;
				try {
					out = new FileOutputStream(
							   new File(SaveStringText.getText())); 
				} catch (FileNotFoundException e1) {
					JOptionPane.showMessageDialog(null, "Error: 2165 Fout opgetreden: " +  e1.getMessage() );;
				}
			

				
	String	kvQuery_c = kv_sl_cont.substring(1, kv_sl_cont.length()-1);
	String	kvQuery_i = kv_sl_invent.substring(1, kv_sl_invent.length()-1);
	

	
	try {
			  	       Class.forName("org.postgresql.Driver");
			  	       c = DriverManager
			  	          .getConnection("jdbc:postgresql://192.168.10.7:5432/sr",
			  	          "stephan", "playfairs");
			  	       c.setAutoCommit(false);
			  	      
			  	   stmtc_alg = c.createStatement();
			  	   stmtc_tree = c.createStatement();
			  	   stmti_alg = c.createStatement();
			  	   stmti_tree = c.createStatement();
			  	   stmtc_Terrain = c.createStatement();
			  	   stmtci = c.createStatement();
			  	   stmttemp = c.createStatement();
			  	   stmtemp1 = c.createStatement();
			  	   temp_alg = c.createStatement();
			  	   stmttot = c.createStatement();
			 
	} catch(Exception x){JOptionPane.showMessageDialog(null, "connection error: " +  x.getMessage() );}	
	
	 	completion = 30;
	   
	   if(completion == 30){
	   stringDisplay0.setText("performing Acrobats");
	   p.add(stringDisplay0, cp);
	   }
	
	try{
			  	 
			  	   
			  	 if( dayTextfield2.isEmpty()  && monthTextfield2.isEmpty() && yearTextfield2.isEmpty()){
			  		 
			  		 ResultSet rsc_alg = stmtc_alg.executeQuery( "SELECT DISTINCT on (kapvak) odk, ploegleider, COALESCE(boomkenner, 'Geen Data') as boomkenner, "+
			  			"	COALESCE(opnemer3, 'Geen Data') as opnemer3,	"+
			  			"	COALESCE(opnemer4, 'Geen Data') as opnemer4, "+
			  			"	COALESCE( to_char(start_opname, 'DD-MM-YYYY'), '9999-09-09') as start_opname, "+
			  			"	COALESCE(opmerkingen, 'Geen Data') as opmerkingen, "+
			  			"	COALESCE(eerdere_hka, 0) as eerdere_hka, "+
			  			"	COALESCE(markatiemethode, 0) as markatiemethode, "+
			  			"	COALESCE(nietbosbouwactiviteiten, 'Geen Data') as nietbosbouwactiviteiten, "+
			  			"	COALESCE(markatiekleur, 'Geen Data') as markatiekleur, "+
			  			"	COALESCE(afbakening, 0) as afbakening, "+
			  			"	COALESCE(kapvak_num, 0) as kapvak_num, "+
			  			"	COALESCE(markatieafstand, 0) as markatieafstand, "+
			  			"	COALESCE(kapvak, 0) as kapvak, "+
			  			"	COALESCE(type_activiteit, 'Geen Data') as type_activiteit, "+
			  			"	COALESCE(type_boom, 'Geen Data') as type_boom, "+
			  			"	COALESCE(schaal_ehka, 0) as schaal_ehka, "+
			  			"	COALESCE(jaar_ehka, 0) as jaar_ehka, "+
			  			"	COALESCE(boommarkatie, 0) as boommarkatie, "+
			  			"	COALESCE(sleepwegmarkatie, 0) as sleepwegmarkatie "+
			  			"	FROM opname.honderd_procent_algemeen  WHERE terrein ="+ terrein_db + " AND kapvak in ( "+ kvQuery_c  +" ) AND start_opname = "+ date_db +"::date; ");
		      
			  	  while ( rsc_alg.next() ) {
	  	   	  	   
		    	  dataodk.addElement(String.format("'%s'", rsc_alg.getString("odk")));
		    	  datapLeider.addElement(rsc_alg.getString("ploegleider"));
		    	  databKenner.addElement(rsc_alg.getString("boomkenner"));
		    	  dataopnemer3.addElement(rsc_alg.getString("opnemer3"));
		    	  dataopnemer4.addElement(rsc_alg.getString("opnemer4"));
		    
		    	  data_date.addElement(rsc_alg.getString("start_opname"));	
		    	  dataopmc.addElement(rsc_alg.getString("opmerkingen"));
		    	  dataehka.addElement(rsc_alg.getString("eerdere_hka"));
		    	  datakapvn.addElement(rsc_alg.getString("kapvak_num"));
		    	  datamarkmethod.addElement(rsc_alg.getString("markatiemethode"));
		    	  datanbosact.addElement(String.valueOf(rsc_alg.getInt("nietbosbouwactiviteiten")));
	  	 	      datamarkkleur.addElement(rsc_alg.getString("markatiekleur"));
	  	 	      dataAfbakening.addElement(rsc_alg.getString("afbakening"));
	  	 	      datamark_afst.addElement(rsc_alg.getString("markatieafstand"));
	  	 	      datakvc.addElement(rsc_alg.getString("kapvak"));
	  	 	      data_ta.addElement(rsc_alg.getString("type_activiteit"));
	  	 	      dataJaarehka.addElement(rsc_alg.getString("jaar_ehka"));
	  	 	      dataSchaalehka.addElement(rsc_alg.getString("schaal_ehka"));
	  	 	      databoommark.addElement(rsc_alg.getString("boommarkatie"));
	  	 	      datasleepmark.addElement(rsc_alg.getString("sleepwegmarkatie"));
	  	    	    	  	 		  }
			  	  rsc_alg.close();
			  	  
			     } else if (!dayTextfield2.isEmpty()  && !monthTextfield2.isEmpty() && !yearTextfield2.isEmpty()){
			    	 
			  		 ResultSet rsc_alg = stmtc_alg.executeQuery( "SELECT DISTINCT on (kapvak) odk, ploegleider, COALESCE(boomkenner, 'Geen Data') as boomkenner, "+
					  			"	COALESCE(opnemer3, 'Geen Data') as opnemer3,	"+
					  			"	COALESCE(opnemer4, 'Geen Data') as opnemer4, "+
					  			"	COALESCE( to_char(start_opname, 'DD-MM-YYYY'), '9999-09-09') as start_opname, "+
					  			"	COALESCE(opmerkingen, 'Geen Data') as opmerkingen, "+
					  			"	COALESCE(eerdere_hka, 0) as eerdere_hka, "+
					  			"	COALESCE(markatiemethode, 0) as markatiemethode, "+
					  			"	COALESCE(nietbosbouwactiviteiten, 'Geen Data') as nietbosbouwactiviteiten, "+
					  			"	COALESCE(markatiekleur, 'Geen Data') as markatiekleur, "+
					  			"	COALESCE(afbakening, 0) as afbakening, "+
					  			"	COALESCE(kapvak_num, 0) as kapvak_num, "+
					  			"	COALESCE(markatieafstand, 0) as markatieafstand, "+
					  			"	COALESCE(kapvak, 0) as kapvak, "+
					  			"	COALESCE(type_activiteit, 'Geen Data') as type_activiteit, "+
					  			"	COALESCE(type_boom, 'Geen Data') as type_boom, "+
					  			"	COALESCE(schaal_ehka, 0) as schaal_ehka, "+
					  			"	COALESCE(jaar_ehka, 0) as jaar_ehka, "+
					  			"	COALESCE(boommarkatie, 0) as boommarkatie, "+
					  			"	COALESCE(sleepwegmarkatie, 0) as sleepwegmarkatie "+
					  			"FROM opname.honderd_procent_algemeen  WHERE terrein ="+ terrein_db + " AND kapvak in ( "+ kvQuery_c  +" ) " +
					  			"AND start_opname >= " + date_db + "::date AND start_opname <= "+ date_db2 +"::date ;" );
		      
			  	  while ( rsc_alg.next() ) {
			  		  dataodk.addElement(String.format("'%s'", rsc_alg.getString("odk")));
			    	  datapLeider.addElement(rsc_alg.getString("ploegleider"));
			    	  databKenner.addElement(rsc_alg.getString("boomkenner"));
			    	  dataopnemer3.addElement(rsc_alg.getString("opnemer3"));
			    	  dataopnemer4.addElement(rsc_alg.getString("opnemer4"));
			    	  data_date.addElement(rsc_alg.getString("start_opname"));	 
			    
			    	  dataopmc.addElement(rsc_alg.getString("opmerkingen"));
			    	  dataehka.addElement(rsc_alg.getString("eerdere_hka"));
			      	  datakapvn.addElement(rsc_alg.getString("kapvak_num"));
			    	  datamarkmethod.addElement(rsc_alg.getString("markatiemethode"));
			    	  datanbosact.addElement(String.valueOf(rsc_alg.getInt("nietbosbouwactiviteiten")));
		  	 	      datamarkkleur.addElement(rsc_alg.getString("markatiekleur"));
		  	 	      dataAfbakening.addElement(rsc_alg.getString("afbakening"));
		  	 	      datamark_afst.addElement(rsc_alg.getString("markatieafstand"));
		  	 	      datakvc.addElement(rsc_alg.getString("kapvak"));
		  	 	      data_ta.addElement(rsc_alg.getString("type_activiteit"));
		  	 	      dataJaarehka.addElement(rsc_alg.getString("jaar_ehka"));
		  	 	      dataSchaalehka.addElement(rsc_alg.getString("schaal_ehka"));
		  	 	      databoommark.addElement(rsc_alg.getString("boommarkatie"));
		  	 	      datasleepmark.addElement(rsc_alg.getString("sleepwegmarkatie"));
	  	    	    	  	 		  }
			  	  rsc_alg.close();
			  	  
			    	 
			     }
		
			       
	} catch(Exception x){JOptionPane.showMessageDialog(null, "Error algemene informatie: " +  x.getMessage() );}	


	
    String dataodk_string = dataodk.toString();
    String dataodk_pg	 = dataodk_string.substring(1, dataodk_string.length()-1);
    
	completion = 40;
	   
	   if(completion == 40){
    stringDisplay0.setText("Executing backflip!");
	   p.add(stringDisplay0, cp);
	   }
    ResultSet totaltrees;
	try {
		totaltrees = stmttot.executeQuery(
				   "SELECT DISTINCT  CONCAT(honderd_procent_boom.kapvak, LPAD(honderd_procent_boom.perceel::text, 2, '0'), honderd_procent_boom.boomnummer) AS sucont,"
				   + " 	CONCAT(invent.kapvak, LPAD(invent.perceel::text, 2, '0'), invent.boomnummer) AS suinv FROM opname.honderd_procent_boom, inventarisatie.invent"
				   + "	WHERE CONCAT(honderd_procent_boom.kapvak, LPAD(honderd_procent_boom.perceel::text, 2, '0'), honderd_procent_boom.boomnummer) = "
				   + "  CONCAT(invent.kapvak, LPAD(invent.perceel::text, 2, '0'), invent.boomnummer) AND"
				   + " 	opname.honderd_procent_boom.formulier in(" + dataodk_pg + ") AND "
				   + "  terrein ="+ terrein_db +";");

	   
	   while ( totaltrees.next()){
		   
		 dataTotalTrees.addElement(totaltrees.getString("sucont"));
		   
	   }
	   
	}  catch(Exception x){JOptionPane.showMessageDialog(null, "Error totaal aantal bomen: " +  x.getMessage() );}	
	   
	completion = 50;
	   
	   if(completion == 50){
	 stringDisplay0.setText("Trying skydiving...");
	   p.add(stringDisplay0);
	   }
	   
	try{
	
		  	temp_alg.execute(
		  			"DELETE FROM jtemps.temptablecont;	"+
		  			"DELETE FROM jtemps.temptableinvent;	"+	
		  			"DELETE FROM jtemps.combtable;	"+
		  			"DROP TABLE IF EXISTS tempfill;	"+

		  			"CREATE TABLE tempfill(		"+
		  			"uid character varying(50),	"+
		  			"kapvak integer,	"+
		  			"bostypecont integer,	"+
		  			"bodemtypecont integer,	"+
		  			"topografiecont integer,	"+
		  			"ondergroei character varying(2000) );	"+

		  						"INSERT INTO tempfill( uid, kapvak,  bostypecont, bodemtypecont, topografiecont, ondergroei)	"+	
		  						"SELECT  parentid , kapvak ,	"+
		  					  	"first_value(nullif(bostype, 0)) OVER (order by parentid, id_position  ROWS UNBOUNDED PRECEDING ) bostype,	"+	
		  					  	"first_value(nullif(bodemtype, 0)) OVER (order by parentid, id_position  ROWS UNBOUNDED PRECEDING ) bodemtype,	"+ 
		  					  	"first_value(nullif(topografie, 0)) OVER (order by parentid, id_position  ROWS UNBOUNDED PRECEDING )  topografie,		"+
		  					  	"first_value(nullif(ondergroei, '0')) OVER (order by parentid, id_position  ROWS UNBOUNDED PRECEDING ) ondergroei		"+
		  						"FROM    opname.plot_info_cont	"+
		  						"WHERE parentid in (" + dataodk_pg +" );	"+

		  					  	"INSERT INTO jtemps.temptablecont( uid, kapvak,  bostypecont, bodemtypecont, topografiecont, ondergroei, bodemomschrijving, bosomschrijving, topoomschrijving)	"+	
		  						"SELECT uid, kapvak,  bostypecont, bodemtypecont, topografiecont, ondergroei,   bodemtype.omschrijving, bostype.omschrijving, topo.omschrijving	FROM    tempfill	"+
		  					  	"LEFT JOIN  lut.bodemtype ON bodemtype.code::integer = tempfill.bodemtypecont::integer	"+
		  					  	"LEFT JOIN  lut.bostype   ON bostype.code::integer = tempfill.bostypecont::integer		"+
		  					  	"LEFT JOIN  lut.topo      ON topo.code::integer = tempfill.topografiecont::integer;		"+
		  					  	
		  						"INSERT INTO jtemps.temptableinvent(kapvak,  bostypeinvent, bosomschrijving, bodemtypeinvent, bodemomschrijving, topografieinvent, topoomschrijving)	"+	
		  					  	"SELECT DISTINCT kapvak, bostype, bostype.omschrijving, bodemtype, bodemtype.omschrijving, topografie, topo.omschrijving		"+
		  					  	"FROM inventarisatie.plot_info	 	"+
		  					  	"LEFT JOIN  lut.bodemtype ON bodemtype.code::integer = plot_info.bodemtype::integer	"+	
		  					  	"LEFT JOIN  lut.bostype   ON bostype.code::integer = plot_info.bostype::integer		"+
		  					  	"LEFT JOIN  lut.topo      ON topo.code::integer = plot_info.topografie::integer		"+
		  					  	"WHERE terrein =" + terrein_db +"AND kapvak in (" + kvQuery_i +  ");	 	"+

		  					  	"INSERT INTO jtemps.combtable ( kapvak, bosinvent, boscont, ondergroei, bodeminvent, bodemcont, topoinvent, topocont)	"+	
		  					  	"SELECT DISTINCT temptablecont.kapvak, string_agg( DISTINCT temptableinvent.bosomschrijving, ', ') bosinvent, string_agg( DISTINCT temptablecont.bosomschrijving, ', ') boscont, ondergroei,	"+ 
		  					  	"string_agg( DISTINCT temptableinvent.bodemomschrijving, ', ') bodeminvent, string_agg(DISTINCT temptablecont.bodemomschrijving, ', ') bodemcont,	 	"+
		  					  	"string_agg( DISTINCT temptableinvent.topoomschrijving, ', ')topoinvent, string_agg(DISTINCT temptablecont.topoomschrijving, ', ') topocont		"+
		  					  	"FROM jtemps.temptablecont, jtemps.temptableinvent WHERE temptablecont.kapvak = temptableinvent.kapvak GROUP BY temptablecont.kapvak, temptablecont.ondergroei;	"	
);
		  	temp_alg.close();
		  	
		  	ResultSet rsi_alg = stmti_alg.executeQuery("SELECT * FROM jtemps.combtable");
		  	
		      while ( rsi_alg.next() ) {
	  	 	  	datakvbtype.addElement( rsi_alg.getString("kapvak"));  	 
	  	 	 	datatopoi.addElement( rsi_alg.getString("topoinvent"));
	  	 	 	databodemi.addElement(rsi_alg.getString("bodeminvent"));
	  	 	 	databosi.addElement(rsi_alg.getString("bosinvent"));
	  	 	 	datatopoc.addElement(rsi_alg.getString("topocont"));
		    	databodemc.addElement(rsi_alg.getString("bodemcont"));
		    	databosc.addElement(rsi_alg.getString("boscont"));
		    	dataondergroei.addElement(rsi_alg.getString("ondergroei"));
	  	   	  	 		  }
		      
		      rsi_alg.close();
	}catch(Exception x){JOptionPane.showMessageDialog(null, "Error milieu gegevens: " +  x.getMessage() );}	
	
	completion = 60;
	   
	   if(completion == 60){
	 stringDisplay0.setText("counting stars.");
	   p.add(stringDisplay0);}
	
	try{
		      ResultSet rsc_tree = stmtc_tree.executeQuery( "SELECT kapvak, perceel, boomnummer, boomnummerna, boomcode, dbh, strh, hc, "
		      												+ "	kwaliteit, stronk, boomnummerna, markatie FROM opname.honderd_procent_boom"
		      												+ " WHERE formulier in (" + dataodk_pg + ") AND kapvak in ("+ kvQuery_c +")"  );
				      
				      while ( rsc_tree.next() ) {
			  	 	  	  	 		 
			  	 	
			  	 	 	datakvctc.addElement(rsc_tree.getString("kapvak"));
			  	 	 	datapercc.addElement(rsc_tree.getString("perceel"));
			  	 	 	databoomc.addElement(rsc_tree.getString("boomnummer"));
		
			  	 	 	datadbhc.addElement(rsc_tree.getString("dbh"));
			  	 	 	datastrhc.addElement(rsc_tree.getString("strh"));
			  	 	 	datahcc.addElement(rsc_tree.getString("hc"));
			  	 	 	datakwac.addElement(rsc_tree.getString("kwaliteit"));
			  	 	 	datamark.addElement(rsc_tree.getString("markatie"));
			  	 	 	dataStronk.addElement(rsc_tree.getString("stronk"));
			  	 	 	databoomna.addElement(rsc_tree.getString("boomnummerna"));
			 
			  	 	
			  	 	 	if (Integer.parseInt(rsc_tree.getString("perceel")) < 10){
			  	 		dataSuc.addElement(String.format("%s%s0%s%s",terreinTextfield, rsc_tree.getString("kapvak"),
			  	 				rsc_tree.getString("perceel"),  rsc_tree.getString("boomnummer")));
			  	 		 
			  	 	 } else {dataSuc.addElement(String.format("%s%s%s%s",terreinTextfield, rsc_tree.getString("kapvak"),
			  	 				rsc_tree.getString("perceel"),  rsc_tree.getString("boomnummer")));
			  	 	 	}
				      }
				      
				      rsc_tree.close();
	} catch(Exception x){JOptionPane.showMessageDialog(null, "Error 1700: " +  x.getMessage() );}	
			
	completion = 70;
	   
	   if(completion == 70){
	 stringDisplay0.setText("Letting the dogs out.");
	   p.add(stringDisplay0);}
	   
	try{
			   ResultSet rsi_tree = stmti_tree.executeQuery( "SELECT samplingunit, kapvak, perceel, boomnummer, houtsoortcode, dbh, hst, hc, "
					   										+"kwaliteit, selectie, inventarisatiedatum, opnemer FROM inventarisatie.invent WHERE terrein =" 
					   										+ terrein_db + "AND kapvak in (" + kvQuery_i +") AND dbh||hst||hc||kwaliteit||selectie IS NOT NULL;" );
						      
						      while ( rsi_tree.next() ) {
					  	 	  	  	 		 
					  	 	 	dataSui.addElement( rsi_tree.getString("samplingunit"));
					  	 	 	datakvcti.addElement(rsi_tree.getString("kapvak"));
					  	 	 	dataperci.addElement(rsi_tree.getString("perceel"));
					  	 	 	databoomi.addElement(rsi_tree.getString("boomnummer"));
					  
					  	 	 	datadbhi.addElement(rsi_tree.getString("dbh"));
					  	 	 	datastrhi.addElement(rsi_tree.getString("hst"));
					  	 	 	datahci.addElement(rsi_tree.getString("hc"));
					  	 	 	datakwai.addElement(rsi_tree.getString("kwaliteit"));
					  	 	 	datasel.addElement(rsi_tree.getString("selectie"));	
	
					  			
					  			if (Integer.parseInt(rsi_tree.getString("perceel")) < 10){
						  	 		dataUIDc.addElement(String.format("%s%s0%s%s",terreinTextfield, rsi_tree.getString("kapvak"),
						  	 				rsi_tree.getString("perceel"),  rsi_tree.getString("boomnummer")));
						  	 		 
						  	 	 } else {dataUIDc.addElement(String.format("%s%s%s%s",terreinTextfield, rsi_tree.getString("kapvak"),
						  	 				rsi_tree.getString("perceel"),  rsi_tree.getString("boomnummer")));
						  	 	 	}
					  		
					  	   	  	 		  }
						      rsi_tree.close();
		   
	} catch(Exception x){JOptionPane.showMessageDialog(null, "boomgegevens " +  x.getMessage() );}	
	
	try{

			  ResultSet verb = stmti_tree.executeQuery( "SELECT  kapvak::int, houtsoortcode, COUNT(houtsoortcode)::int FROM  inventarisatie.invent "
			  		+ "WHERE  houtsoortcode in ('BOL', 'TON', 'INT', 'MRZ', 'ROZ', 'SAW', 'HPH') "
			  		+ "AND terrein =" + terrein_db + "AND kapvak in (" + kvQuery_i +") AND selectie = '1' GROUP BY kapvak, houtsoortcode ORDER BY kapvak;");

										      
			  
			  while(verb.next() ){
				  
				  dataKv_verb.addElement(verb.getInt("kapvak"));
				  dataHs_verb.addElement(verb.getString("houtsoortcode"));
				  dataNum_verb.addElement(verb.getInt("count"));
				  			  }
			  verb.close();
	} catch(Exception x){JOptionPane.showMessageDialog(null, "Error selectie verboden soorten: " +  x.getMessage() );}	
	
	 try{
			  ResultSet onderma = stmti_tree.executeQuery("SELECT kapvak::int, houtsoortcode, COUNT(houtsoortcode) FROM "
			  		+ " inventarisatie.invent  WHERE terrein =" + terrein_db + "AND kapvak in (" + kvQuery_i +
			  			") AND selectie = '1' AND dbh::double precision < 35 GROUP BY kapvak::int, houtsoortcode  ORDER BY kapvak;" );

		      

while(onderma.next() ){

dataKv_onderma.addElement(onderma.getInt("kapvak"));
dataHs_onderma.addElement(onderma.getString("houtsoortcode"));
dataNum_onderma.addElement(onderma.getInt("count"));

}

onderma.close();
	 }  catch(Exception x){JOptionPane.showMessageDialog(null, "Error selectie ondermaatse soorten: " +  x.getMessage() );}	
	 
	 
	 try{
ResultSet inventdat = stmti_tree.executeQuery(
"SELECT kapvak, CASE WHEN NOT min(inventdatum::date) = max(inventdatum::date) THEN CONCAT(min(to_char(inventdatum::date, 'DD-MM-YYYY')) , ' t/m ', "+
"max(to_char(inventdatum::date, 'DD-MM-YYYY'))) ELSE max(to_char(inventdatum::date, 'DD-MM-YYYY')) END as inventdatum from inventarisatie.plot_info  where kapvak in ("
		+ kvQuery_i + ") AND terrein =" + terrein_db + " GROUP BY kapvak;" );

while(inventdat.next() ){

	datainventdati.addElement(inventdat.getString("inventdatum"));

}

inventdat.close();

	 } catch(Exception x){JOptionPane.showMessageDialog(null, "Error PG inventarisatie datum: " +  x.getMessage() );}	

	 
	 try{
ResultSet rs_opnemer = stmtci.executeQuery("SELECT DISTINCT kapvak, string_agg(DISTINCT COALESCE(opnemer, 'Geen Data'), ', ') opnemer from inventarisatie.invent"
		+  " where kapvak in ("+  kvQuery_i + ") AND terrein = " + terrein_db + " GROUP BY kapvak ORDER BY kapvak; " );

while(rs_opnemer.next()){
datakvi.addElement(rs_opnemer.getString("kapvak"));
dataopnemeri.addElement(rs_opnemer.getString("opnemer"));


}
rs_opnemer.close();
	 } catch(Exception x){JOptionPane.showMessageDialog(null, "Error PG opnemer: " +  x.getMessage() );}	

	 completion = 80;
	   
	   if(completion == 80){
	 stringDisplay0.setText("doowakkies enabled and functioning within normal parameters.");
	   p.add(stringDisplay0);
	   }
	   
try{

for (int i = 0; i < numKvSelect; i ++ ){

	stmttemp.execute(
			"DELETE FROM jtemps.temptable; 	"+
			"DELETE FROM jtemps.tempselecttable;	"+

			"INSERT INTO jtemps.tempselecttable(formulier, kapvak, terrein, boomnummer, markatie, selectie)	"+ 
			 
			"SELECT  formulier, inventarisatie.invent.kapvak, terrein,  opname.honderd_procent_boom.boomnummer, markatie, selectie::int 	"+
								
			"FROM 	opname.honderd_procent_boom, inventarisatie.invent	"+

				"WHERE formulier in ("+ dataodk_pg +")	"+
				"AND   inventarisatie.invent.kapvak::int = opname.honderd_procent_boom.kapvak::int	"+
				"AND   opname.honderd_procent_boom.boomnummer::int  = inventarisatie.invent.boomnummer::int "+
				"AND   inventarisatie.invent.perceel::int = opname.honderd_procent_boom.perceel::int 	" +
				"AND   inventarisatie.invent.terrein =" + terrein_db +	
				"AND   inventarisatie.invent.kapvak in (" + datakvi.get(i) + ");	"+

				"INSERT INTO jtemps.temptable (kapvak, column1, column2, column3, column4) "+	
				"SELECT kapvak, 0,0,0,0 FROM jtemps.tempselecttable GROUP BY kapvak; "+

				"UPDATE jtemps.temptable	"+	
				"set column1 = aantal	"+
				"FROM(	"+	
				"SELECT  kapvak, COUNT(kapvak) as aantal	"+	 
				"FROM 	jtemps.tempselecttable	"+	
				"WHERE 	markatie::int = 1	"+	
				"AND   selectie::int = 1	"+	
				"GROUP BY  kapvak) AS aantal;	"+
				
			"UPDATE jtemps.temptable	"+
			"set column2 = aantal	"+
			"FROM( SELECT  kapvak, COUNT(kapvak) as aantal	"+ 
				"FROM 	jtemps.tempselecttable	"+
				"WHERE   NOT markatie::int = 1	"+
				"AND   selectie::int = 1	"+

			"GROUP BY  kapvak) AS aantal;	"+

			"UPDATE jtemps.temptable	"+
			"set column3 = aantal	"+
			"FROM( SELECT  kapvak, COUNT(kapvak) as aantal	"+ 
				"FROM 	jtemps.tempselecttable	"+
				"WHERE   markatie::int = 1	"+
				"AND   NOT selectie::int = 1	"+

			"GROUP BY  kapvak) AS aantal;	"+

			"UPDATE jtemps.temptable	"+
			"set column4 = aantal	"+
			"FROM( SELECT  kapvak, COUNT(kapvak) as aantal	"+ 
			"FROM 	jtemps.tempselecttable	"+
			"WHERE NOT  markatie::int = 1	"+
			"AND NOT  selectie::int = 1		"+
			"GROUP BY  kapvak) AS aantal;	"
 );	
	
					

	
					ResultSet rs = stmtci.executeQuery("SELECT * from jtemps.temptable;");
					
					while(rs.next()){
						
						markselKV.addElement(rs.getInt("kapvak"));
						mark1sel1.addElement(rs.getInt("column1"));
						mark0sel1.addElement(rs.getInt("column2"));
						mark1sel0.addElement(rs.getInt("column3"));
						mark0sel0.addElement(rs.getInt("column4"));
					
									
					}
						
					rs.close();	
}		
stmttemp.close();
} catch(Exception x){JOptionPane.showMessageDialog(null, "Error markatie v selectie: " +  x.getMessage() );}			

try{

					ResultSet rs_boomcodes_goed = stmtci.executeQuery(" SELECT DISTINCT honderd_procent_boom.kapvak, honderd_procent_boom.boomnummer, honderd_procent_boom.boomcode, invent.houtsoortcode "+
																 	  " FROM opname.honderd_procent_boom JOIN  inventarisatie.invent ON  invent.boomnummer::text = honderd_procent_boom.boomnummer::text AND " +
																      " invent.kapvak = honderd_procent_boom.kapvak AND invent.perceel::text = honderd_procent_boom.perceel::text AND "+
																      " honderd_procent_boom.boomcode  = houtsoortcode WHERE invent.terrein ="+ terrein_db +"AND "+
																      " honderd_procent_boom.formulier in ( "+ dataodk_pg +" );");
 
					while(rs_boomcodes_goed.next()){
						datahscgoed.addElement(rs_boomcodes_goed.getString("boomcode"));
						}
					
										
					ResultSet rs_boomcodes_klfout = stmtci.executeQuery(
							"WITH controle as ( " +
							"SELECT DISTINCT  opname.honderd_procent_boom.boomcode, "+
							"jtemps.foutenlijst.id, "+
							"opname.honderd_procent_boom.kapvak, "+
							"opname.honderd_procent_boom.perceel, "+
							"opname.honderd_procent_boom.boomnummer "+	
											 
							"FROM opname.honderd_procent_boom, jtemps.foutenlijst "+ 	
							"WHERE jtemps.foutenlijst.code = opname.honderd_procent_boom.boomcode AND "+	
							"opname.honderd_procent_boom.formulier in (" + dataodk_pg + ") ),	"+
							"inventarisatie as ( "+

							"SELECT DISTINCT  inventarisatie.invent.houtsoortcode, "+	
									  		 "jtemps.foutenlijst.id, "+ 
									  		 "inventarisatie.invent.kapvak, "+
									  		 "inventarisatie.invent.perceel, "+
									  		 "inventarisatie.invent.boomnummer	"+

											"FROM inventarisatie.invent, jtemps.foutenlijst "+	
											"WHERE   jtemps.foutenlijst.code = inventarisatie.invent.houtsoortcode AND "+	
												"inventarisatie.invent.terrein = " + terrein_db + " AND	"+
							"inventarisatie.invent.kapvak in ("+ kvQuery_i  +")) "+

							"SELECT  controle.id, inventarisatie.houtsoortcode, controle.boomnummer, inventarisatie.boomnummer FROM "+	
										"controle, inventarisatie "+	
										"WHERE controle.id = inventarisatie.id AND "+ 
										"controle.boomnummer::integer = inventarisatie.boomnummer::integer AND "+
										"controle.kapvak::integer = inventarisatie.kapvak::integer AND "+
										"controle.perceel::integer = inventarisatie.perceel::integer AND "+	
										"inventarisatie.houtsoortcode != controle.boomcode;"
							);
					
					
					
					
					while(rs_boomcodes_klfout.next()){
						datahscklfout.addElement(rs_boomcodes_klfout.getString("id"));
						}
					
					  rs_boomcodes_goed.close();
					  rs_boomcodes_klfout.close();
					  
					  
		} catch (Exception klfout){JOptionPane.showMessageDialog(null, "Error houtsoort vergelijking opgetreden: " +  klfout.getMessage() );}	
		  
					 
					try{   
					   stmtc_alg.close(); 
				  	   stmtc_tree.close();
				  	   stmti_alg.close(); 
				  	   stmti_tree.close();
				  	   stmtc_Terrain.close(); 
				  	   stmtci.close(); 
				  	   stmttemp.close();
				  	   stmtemp1.close();
				  	   temp_alg.close();
			           c.close();
			           
					} catch (Exception n){JOptionPane.showMessageDialog(null, "Error closing statements: " +  n.getMessage() );}	
			        
 //  } catch (Exception PG1){JOptionPane.showMessageDialog(null, "Error Postgress opgetreden: " +  PG1.getMessage() );}	    
	 
					completion = 90;
					   
					   if(completion == 90){
					   stringDisplay0.setText("deploying monkeys!");
					   p.add(stringDisplay0);}
	 
	 countboomcodes = datahscgoed.size() + datahscklfout.size() ;
	 
	 double doublecountboomcodes =  datahscgoed.size() + datahscklfout.size() ;
	 
	 double totboomcodes = 	dataTotalTrees.size();
	 
	 double countboomcodesproc = (doublecountboomcodes/totboomcodes)*100;
	 
	 int foutboom = dataTotalTrees.size() - countboomcodes;
	 
	 double doublefoutboom = foutboom;
			 
	 double foutboomproc = (doublefoutboom/totboomcodes)*100;
	 

	 
     for (int i = 0; i < datamarkmethod.size(); i++){
 String localVar = datamarkmethod.get(i);
 
   switch(localVar){
   
   case "1": 
  	  	 datamarkmethodString.addElement("Verf");
   break;	
   
   case "2":
  	 	datamarkmethodString.addElement("Spuitbom");
  break;
  
   case "3":
  	 	datamarkmethodString.addElement("Flagging");
  break;
  
   case "0":
  	 	datamarkmethodString.addElement("Niet gemarkeerd");
  break;
     }
};
     
     
     
     for( int i = 0; i < datasleepmark.size(); i++ ){
  	   
  	   String localVar = datasleepmark.get(i);
     switch(localVar){
  
     case "1":
  	   datasleepmarkString.addElement("Ja");
  
     break;
     
     case "2":
  	   datasleepmarkString.addElement("Nee");
  	break;
     
     }
  };
     
     for( int i = 0; i < databoommark.size(); i ++ ){
  	   
  	   String localVar = databoommark.get(i);
  	   
	       switch(localVar){
	    
	       case "1":
	    	   databoommarkString.addElement("Ja");
	    
	       break;
	       
	       case "2":
	    	   databoommarkString.addElement("Nee");
  	    break;
	       }
	    };
     
   
	    
     for( int i = 0; i <  dataJaarehka.size(); i++ ){
  	   
    	 String localVar = null;
    	 
    	 if (dataJaarehka.get(i).equals("null")){
    		 
    	 localVar = "NA";} else {
    	 
  	     localVar = dataJaarehka.get(i);}
  	   
	       switch (localVar){
	    
	       case "1":
	    	   datamarkJaarString.addElement("Minder dan 1 jaar");
	    
	       break;
	       
	       case "2":
	    	   datamarkJaarString.addElement("1 tot 5 jaar");
	       
	       break;
	       
	       case "3":
	    	   datamarkJaarString.addElement("Meer dan 5 jaar");
	       break;
	       
	       default: datamarkJaarString.addElement("Geen eerdere Houtkap Activiteiten");
	       }
	    };
     
     
     for( int i = 0; i <  dataSchaalehka.size(); i++ ){
  	   
  	   String localVar = dataSchaalehka.get(i);
  	   
	       switch (localVar){
	    
	       case "1":
	    	   datamarkSchaalString.addElement("Kleinschalig");
	    
	       break;
	       
	       case "2":
	    	   datamarkSchaalString.addElement("Grootschalig");
	       break;  
	       
	       default: datamarkSchaalString.addElement("Geen eerdere Houtkap Activiteiten");
	       }
	    };
     

     
     for( int i = 0; i < datanbosact.size(); i++ ){
  	   
  	   if (datanbosact.get(i).equals("2")){
	     
	    	nbosactString = "Nee";
	    	  
	    			    	   } else {nbosactString = "Ja";   
	    			    	   break; }	
     };
	 
	       
	    
     
     for(int i = 0; i < dataehka.size(); i++){
  	 
  	 if(dataehka.get(i).equals("2")){
  	    
  	 	ehkaString =  "Nee"; 
         
        } else {ehkaString = "Ja";
        break;}
     };

     for(int i = 0; i < datakapvn.size(); i++){
  	   
  	   String localVar = datakapvn.get(i);
  	   
  	   switch( localVar){
  	      	   
  	   case "1": datakapvnString.addElement("Goed");
  	   
  	   break;
  	   
  	   case "2": datakapvnString.addElement("Matig");
  	   
  	   break;
  	   
  	   case "3": datakapvnString.addElement("Slecht");
  	   
  	   break;
  	   }
     };
  	
     for(int i = 0; i < dataAfbakening.size(); i++){
  	   
  	   String localVar = dataAfbakening.get(i);
  	   
  	   switch( localVar ){
  	   
  	   case "1": dataAfbString.addElement("Goed");
  	   
  	   break;
  	   
  	   case "2": dataAfbString.addElement("Matig");
  	   
  	   break;
  	   
  	   case "3": dataAfbString.addElement("Slecht");
  	   
  	   break;
  	   }
     };
  	   	 
        uniqueDate.addAll( data_date);
		uniquepLeiders.addAll(datapLeider);
		uniqueBoomKenners.addAll(databKenner);
		uniqueOpn1.addAll(dataopnemer3); 
		uniqueOpn2.addAll(dataopnemer4);
		String uniquepLeiderSize = uniquepLeiders.toString();
		String uniqueBoomKennersSize = uniqueBoomKenners.toString();
		String uniqueOpn1Size = uniqueOpn1.toString();
		String uniqueOpn2Size = uniqueOpn2.toString();

		
		String	pLeidersString = uniquepLeiders.toString().substring(1, uniquepLeiderSize.length()-1); 
		
		String	boomkennersString = uniqueBoomKenners.toString().substring(1, uniqueBoomKennersSize.length()-1); 
		String	opnemers1 = uniqueOpn1.toString().substring(1, uniqueOpn1Size.length()-1); 
		String	opnemers2 = uniqueOpn2.toString().substring(1, uniqueOpn2Size.length()-1); 
	 
	 //Markatie 
	 Vector<String> mark1Temp = new Vector<String>();
	 Vector<String> mark2Temp = new Vector<String>();
	 Vector<String> mark3Temp = new Vector<String>();
	 Vector<String> mark4Temp = new Vector<String>();
	 Vector<String> mark5Temp = new Vector<String>();
	 Vector<String> mark0Temp = new Vector<String>();
	 
	 for (int i = 0; i < datamark.size(); i++){
		 
		 		if(datamark.get(i).equals("1") ){ 
		 
			 mark1Temp.addElement(datamark.get(i));	 
		 
		 }else if (datamark.get(i).equals("2") ){ 
		 
			 mark2Temp.addElement(datamark.get(i));	 
		 
		 }else if (datamark.get(i).equals("3") ){ 
		 
			 mark3Temp.addElement(datamark.get(i));	 
		 
		 }else if (datamark.get(i).equals("4") ){ 
		 
			 mark4Temp.addElement(datamark.get(i));	 
		 
		 }else if (datamark.get(i).equals("5") ){ 
		 
			 mark5Temp.addElement(datamark.get(i));	 
		 
		 }else if (datamark.get(i).equals("0") ){ 
		 
			 mark0Temp.addElement(datamark.get(i));	 
		 
		 }
		 		 
}
	 int mark0 = mark0Temp.size();
	 int mark1 = mark1Temp.size();
	 int mark2 = mark2Temp.size(); 
	 int mark3 = mark3Temp.size();
	 int mark4 = mark4Temp.size();
	 int mark5 = mark5Temp.size();
	 
	 //Einde Markatie
	 
	 completion = 95;
	   
	   if(completion == 95){
	 stringDisplay0.setText("Releasing hounds!");
	   p.add(stringDisplay0);
	   }
	   
	 // DBH en HC vergelijking
	
	 for (int i = 0; i < dataSuc.size(); i ++){
	 		
	 		 for(int j = 0; j < dataUIDc.size(); j ++){
	 						 
	 			 if(dataSuc.get(i).equals(dataUIDc.get(j))) {
			 
	 				 dataDeltadbh.addElement(
	 						 Integer.parseInt(datadbhc.get(i)) - Integer.parseInt(datadbhi.get(j)) );
	 				 
	 				 dataDeltahc.addElement(
	 						 Integer.parseInt(datahcc.get(i)) - Integer.parseInt(datahci.get(j)) );
	 				 
	 			 }
	 		 }	 
	 	 }

	 Vector<Integer> tempClassDbh1 = new Vector<Integer>();
	 Vector<Integer> tempClassDbh2 = new Vector<Integer>();
	 Vector<Integer> tempClassDbh3 = new Vector<Integer>();
	 
	 for (int i = 0; i < dataDeltadbh.size(); i ++ ){
		 
 		
 		 
		 if( dataDeltadbh.get(i) <= 5 ){
			 tempClassDbh1.addElement(dataDeltadbh.get(i));
			 
		 } else if ( dataDeltadbh.get(i) > 5 && dataDeltadbh.get(i) <= 10 ){
			 tempClassDbh2.addElement(dataDeltadbh.get(i));
		 
		 } else if ( dataDeltadbh.get(i) > 10 ){
			 tempClassDbh3.addElement(dataDeltadbh.get(i));
			 
		 }
	 }	
		int sizeDbhClass1 = tempClassDbh1.size();
		double sizeDbhClass1_db = tempClassDbh1.size();
		double sizeDbhClass1_prc = sizeDbhClass1_db/dataDeltadbh.size() * 100;
				
		int sizeDbhClass2 = tempClassDbh2.size();
		double sizeDbhClass2_db = tempClassDbh2.size();
		double sizeDbhClass2_prc = sizeDbhClass2_db/dataDeltadbh.size() * 100;
		
		int sizeDbhClass3 = tempClassDbh3.size();
		double sizeDbhClass3_db = tempClassDbh3.size();
		double sizeDbhClass3_prc = sizeDbhClass3_db/dataDeltadbh.size() * 100;
		 
		
		 Vector<Integer> tempClassHc1 = new Vector<Integer>();
		 Vector<Integer> tempClassHc2 = new Vector<Integer>();
		 Vector<Integer> tempClassHc3 = new Vector<Integer>();
		 
		 for (int i = 0; i < dataDeltahc.size(); i ++ ){
			
			 if( dataDeltahc.get(i) <= 2 ){
				 tempClassHc1.addElement(dataDeltahc.get(i));
				 
			 } else if ( dataDeltahc.get(i) > 2 && dataDeltahc.get(i) <= 5 ){
				 tempClassHc2.addElement(dataDeltahc.get(i));
			 
			 } else if ( dataDeltahc.get(i) > 5 ){
				 tempClassHc3.addElement(dataDeltahc.get(i));
				 
			 }
		 }	
			int sizeHcClass1 = tempClassHc1.size();	
			double sizeHcClass1_db =  tempClassHc1.size();	
			double sizeHcClass1_prc = sizeHcClass1_db/dataDeltahc.size() * 100;
			
			int sizeHcClass2 = tempClassHc2.size();
			double sizeHcClass2_db =  tempClassHc2.size();	
			double sizeHcClass2_prc = sizeHcClass2_db/dataDeltahc.size() * 100;
			
			int sizeHcClass3 = tempClassHc3.size();
			double sizeHcClass3_db =  tempClassHc3.size();	
			double sizeHcClass3_prc = sizeHcClass3_db/dataDeltahc.size() * 100;
			
// Einde DBH en HC vergelijking

	 
// Volume vergelijking
	
	final Vector<Integer> KV_sel1 = new Vector<Integer>();
	final Vector<Integer> KV_sel2 = new Vector<Integer>();
	final Vector<Integer> KV_sel3 = new Vector<Integer>();
	
	HashSet<Integer> ukvsel1 = new HashSet<Integer>();
	HashSet<Integer> ukvsel2 = new HashSet<Integer>();
	HashSet<Integer> ukvsel3 = new HashSet<Integer>();

	final Vector<Double> volArraySel1 = new Vector<Double>();
	final Vector<Double> volArraySel2 = new Vector<Double>();
	final Vector<Double> volArraySel3 = new Vector<Double>();
	

	
	 for (int i = 0;  i < dataUIDc.size(); i ++ ){
		 
		 			if(datasel.get(i).equals("1") && !datadbhi.get(i).equals("")){
		 				
		 				double dbh = Double.parseDouble(datadbhi.get(i));
		 				double height = Double.parseDouble(datahci.get(i));
		 				double str = Double.parseDouble(datastrhi.get(i));
		 				
		 		KV_sel1.addElement(Integer.parseInt(datakvcti.get(i)));
		 		
			 	cub1.addElement(0.75 * (
			 			Math.PI	* Math.pow( (dbh/200), 2  ) *
			 			(height - str)	)	);
			 			
			 	} else  if(datasel.get(i).equals("2") && !datadbhi.get(i).equals("")){
		 		
		 				double dbh = Double.parseDouble(datadbhi.get(i));
		 				double height = Double.parseDouble(datahci.get(i));
		 				double str = Double.parseDouble(datastrhi.get(i));
		 				
				 KV_sel2.addElement(Integer.parseInt(datakvcti.get(i)));
 				
 				cub2.addElement(0.75 * (
 						Math.PI	* Math.pow( (dbh/200), 2  ) *
 						(height - str)	)	);
 			
 				
		 	} else  if(datasel.get(i).equals("3") && !datadbhi.get(i).equals("")){
		 		
		 				double dbh = Double.parseDouble(datadbhi.get(i));
		 				double height = Double.parseDouble(datahci.get(i));
		 				double str = Double.parseDouble(datastrhi.get(i));
		 				
				 KV_sel3.addElement(Integer.parseInt(datakvcti.get(i)));
 				
 				cub3.addElement(0.75 * (
 						Math.PI	* Math.pow( (dbh/200), 2  ) *
 						(height - str)	)	);
			} 
	 };
	
	 ukvsel1.addAll(KV_sel1);
	 ukvsel2.addAll(KV_sel2);
	 ukvsel3.addAll(KV_sel3);
	 
		
	 Object[] kvid1 = ukvsel1.toArray();
	 Object[] kvid2 = ukvsel2.toArray();
	 Object[] kvid3 = ukvsel3.toArray();

	 for (int  i = 0 ; i < kvid1.length; i++){
		 Double sumTemp = 0.0;
		 int kvOpp = Integer.parseInt(kvOppTF[i].getText());
		 
	 	 for (int  j = 0 ; j < KV_sel1.size(); j++){
	 		 
	 	 			 if ( kvid1[i].equals(KV_sel1.get(j) ) ) {
	 	 				  

	 	 				 Vector<Double> temp = new Vector<Double>();
	 	 				 temp.addElement(cub1.get(j));
	 	 				
	 	 				 	
	 	 				 for(double  k : temp){
	 	 					 sumTemp += k;
	 	 				 } 	 			
	 	 			 }
	 	 	 
	 	}
	 	volArraySel1.addElement(sumTemp/kvOpp);
	 }
	 

	 
	 for (int  i = 0 ; i < kvid2.length; i++){
		 Double sumTemp = 0.0;
		 int kvOpp = Integer.parseInt(kvOppTF[i].getText());
	 	 for (int  j = 0 ; j < KV_sel2.size(); j++){
	 		 
	 	 			 if ( kvid2[i].equals(KV_sel2.get(j) ) ) {
	 	 				  

	 	 				 Vector<Double> temp = new Vector<Double>();
	 	 				 temp.addElement(cub2.get(j));
	 	 				
	 	 				 for(double  k : temp){
	 	 					sumTemp += k;
	 	 				 } 	 				 
	 	 			 }
	 	 
	 	}
	 	volArraySel2.addElement(sumTemp/kvOpp);
	 }
	
	 
	 for (int  i = 0 ; i < kvid3.length; i++){
		 Double sumTemp = 0.0;
		 int kvOpp = Integer.parseInt(kvOppTF[i].getText());
	 	 for (int  j = 0 ; j < KV_sel3.size(); j++){
	 		 
	 	 			 if ( kvid3[i].equals(KV_sel3.get(j) ) ) {
	 	 				  

	 	 				 Vector<Double> temp = new Vector<Double>();
	 	 				 temp.addElement(cub3.get(j));
	 	 				
	 	 				  for(double  k : temp){
	 	 					 sumTemp += k;
	 	 				 } 
	 	 		   }
	 	 	
		 	}
	 	volArraySel3.addElement(sumTemp/kvOpp);
	 }

		//BUILD OUTPUT ERROR IN BOX 

//end volume vergelijking
	 
	 // boomcode vergelijking
	 

	 
// Stronken
Vector<String> stronk1 = new Vector<String>();
Vector<String> stronk2 = new Vector<String>();
Vector<String> stronk3 = new Vector<String>();

	try{
for(int i = 0; i < dataStronk.size(); i++ ){
	
		  if(dataStronk.get(i).equals("1")){stronk1.addElement(dataStronk.get(i));
		
	}else if(dataStronk.get(i).equals("2")){stronk2.addElement(dataStronk.get(i));
	
	}else if(dataStronk.get(i).equals("3")){stronk3.addElement(dataStronk.get(i));}

}
	} catch(Exception Stronk){System.out.print("geen stronken");}
int nStronken1 = stronk1.size();
int nStronken2 = stronk2.size();
int nStronken3 = stronk3.size();

//bomenNA
Vector<String> boomna = new Vector<String>();
Vector<String> boomValid = new Vector<String>();


try {
for ( int i = 0; i < databoomc.size(); i++){
	   if (databoomc.get(i).equals("0")){boomna.addElement(databoomc.get(i));
	   }else {boomValid.addElement(databoomc.get(i));} }

} catch (Exception nt){ JOptionPane.showMessageDialog(null, "Error:2107 (Geen boom info) Fout opgetreden: " + nt.getMessage() );}


	   double nboomna =  boomna.size();

	   double totBoom = databoomc.size();

	   double procentboomna =  nboomna / totBoom * 100 ;
	   double procentboomValid = 100 - procentboomna;
	      
	   int nconbomen =  boomValid.size() - dataTotalTrees.size();
//Exceptions
	   
	   if(pLeidersString.isEmpty()){ pLeidersString = "Geen ploegleider informatie"; }
	   if(boomkennersString.isEmpty()){boomkennersString = "Geen boomkenner informatie";}
	   if(opnemers1.isEmpty()){opnemers1 ="Geen opnemer informatie";}
	   if(opnemers2.isEmpty()){opnemers2 = "Geen opnemer informatie";}
	   if(textField_bewerker.getText().isEmpty()){textField_bewerker.setText("Geen Informatie");}
	   if(textField_reg.getText().isEmpty()){textField_reg.setText("Geen Informatie");}
	   if(textField_gps.getText().isEmpty()){textField_gps.setText("Geen Informatie");}
	   if(textField_co.getText().isEmpty()){textField_co.setText("Geen Informatie");}
	   
	/*   for(int i = 0;  i < indieningsDatum.length; i++){
	   if(indieningsDatum[i].isEmpty() ){indieningsDatum[i] = "Geen Informatie"; }
	   }
	   */
	   for(int i = 0;  i <  kvQuery_i.length(); i++){
	   if( dataopnemeri.isEmpty() ){ dataopnemeri.addElement("Geen Informatie");}
	   }
       

// end Exceptions	 
	   
completion = 100;
	   
	   if(completion == 100){
		 stringDisplay0.setText("Building Rapport 100% Inventarisatie Controle.");
		 p.add(stringDisplay0);
	   }
	   
try{	   

	 XWPFDocument document= new XWPFDocument();


	 CTDocument1 doc = document.getDocument();
	 CTBody body = doc.getBody();

	 if (!body.isSetSectPr()) {
	      body.addNewSectPr();
	 }
	 CTSectPr section = body.getSectPr();

	 if(!section.isSetPgSz()) {
	     section.addNewPgSz();
	 }
	 CTPageSz pageSize = section.getPgSz();

	 pageSize.setOrient(STPageOrientation.LANDSCAPE);

	 pageSize.setW(BigInteger.valueOf(16840));
	 pageSize.setH(BigInteger.valueOf(11900));
	 


	XWPFParagraph paragraphTitle = document.createParagraph();
	XWPFRun titleRun = paragraphTitle.createRun();
	titleRun.setText("RAPPORT 100%-INVENTARISATIE CONTROLE");
	titleRun.setFontSize(14);
	titleRun.setBold(true);
	paragraphTitle.setAlignment(ParagraphAlignment.CENTER);
	paragraphTitle.setBorderBottom(Borders.THICK);
	paragraphTitle.setBorderTop(Borders.BASIC_THIN_LINES);
	paragraphTitle.setBorderLeft(Borders.BASIC_THIN_LINES);
	paragraphTitle.setBorderRight(Borders.BASIC_THIN_LINES);

	XWPFTable table = document.createTable();	
	XWPFTableRow tableRowOne = table.getRow(0);
	 table.createRow();	 
	 XWPFTableRow tableRowTwo = table.getRow(1);
	 table.createRow();
	 XWPFTableRow tableRowThree = table.getRow(2);
	 table.createRow();
	 XWPFTableRow tableRowFour = table.getRow(3);
	 table.createRow();
	 XWPFTableRow tableRowFive = table.getRow(4);
	 
 Object[] uniqueDateArray = uniqueDate.toArray();
	
	tableRowOne.createCell();
	tableRowOne.getCell(0).setText("Rapport gemaakt door: " + textField_rap.getText());
	tableRowOne.getCell(1).setText("Terreinno: " + Terrein.getText() );
	if (uniqueDate.size() < 2){
	tableRowTwo.getCell(0).setText("Datum veldwerk: " + uniqueDateArray[0].toString() );
	}else { tableRowTwo.getCell(0).setText("Datum veldwerk: " + uniqueDateArray[0] + " t/m " + uniqueDateArray[uniqueDateArray.length - 1] ); }
	tableRowTwo.createCell();
	tableRowTwo.getCell(1).setText("Vergunning houder: " + textField_bewerker.getText() );
			XWPFTableCell Cell1 = tableRowThree.getCell(0);
		 	XWPFParagraph paragraphSBBWerkers = Cell1.getParagraphArray(0);
		 	XWPFRun SBBWerkers1 = paragraphSBBWerkers.createRun();
			SBBWerkers1.setText("SBB veldmedewerkers: " + String.format("%s, %s,", pLeidersString, boomkennersString));
		 	XWPFParagraph paragraphSBBWerkers2 = Cell1.addParagraph();
			XWPFRun SBBWerkers2 = paragraphSBBWerkers2.createRun();
			SBBWerkers2.setText(String.format("%s, %s", opnemers1, opnemers2 ));
			
	tableRowThree.createCell();
	tableRowThree.getCell(1).setText("Kapvakken: " + kvQuery_i);
	tableRowFour.getCell(0).setText("GPS ingeleverd:  " + textField_dat.getText() );
	tableRowFour.createCell();
	tableRowFour.getCell(1).setText("Regio: " + textField_reg.getText());
	tableRowFive.getCell(0).setText("Door: " + textField_gps.getText());
	tableRowFive.createCell();
	tableRowFive.getCell(1).setText("Coördinator: " + textField_co.getText());

	table.setCellMargins(100, 100, 100, 3000);
	
	
	setTableAlignment(table, STJc.CENTER);
	
	XWPFParagraph paragraph = document.createParagraph();
	XWPFRun run = paragraph.createRun();
	run.addBreak();
	run.addBreak();
	
	XWPFParagraph paragraphTitleKapvakken = document.createParagraph();
	XWPFRun titleRunkapvakken = paragraphTitleKapvakken.createRun();
	titleRunkapvakken.setText("ALGEMENE INFO KAPVAK(KEN)");
	titleRunkapvakken.setFontSize(12);
	titleRunkapvakken.setBold(true);
	paragraphTitleKapvakken.setAlignment(ParagraphAlignment.CENTER);
	paragraphTitleKapvakken.setSpacingAfter(10);
	paragraphTitleKapvakken.setSpacingBefore(100);
	
	XWPFTable table2 = document.createTable();
	XWPFTableRow table2RowOne = table2.getRow(0);
	table2RowOne.getCell(0).setText("Kapvak");
	table2RowOne.createCell().setText("Inventarisatie ploeg: " );
	table2RowOne.createCell().setText("Inventarisatie datum: " );
	table2RowOne.createCell().setText("Indieningsdatum: ");
	table2RowOne.createCell().setText("Sleepwegen gepland: " );
	table2RowOne.createCell().setText("GIS data ingediend: ");
	table2RowOne.createCell().setText("Historische Productie: ");
	
	for(int i = 0; i < numKvSelect; i ++  ){
		table2.createRow();
		XWPFTableRow table2Row = table2.getRow(1+i);
		
		table2Row.getCell(0).setText(datakvi.get(i));
		
		table2Row.getCell(1).setText(dataopnemeri.get(i));
		
		table2Row.getCell(2).setText(datainventdati.get(i));
		
		table2Row.getCell(3).setText(indieningsDatum.get(i));
		
		if(sleepweg[i].isSelected())
		{table2Row.getCell(4).setText("Ja");
		}else {table2Row.getCell(4).setText("Nee");}
		
		if(GIS[i].isSelected())
		{	table2Row.getCell(5).setText("Ja");
		}else {table2Row.getCell(5).setText("Nee");}
		
		if(hisProd[i].isSelected())
		{	table2Row.getCell(6).setText("Ja");
		}else {table2Row.getCell(6).setText("Nee");}				
	}
	
	table2.setWidth(2);
	table2.setCellMargins(50,100, 50, 100);
	
	setTableAlignment(table2, STJc.CENTER);
	
	XWPFParagraph paragraph1 = document.createParagraph();
	XWPFRun run1 = paragraph1.createRun();
	run1.addBreak();
	run1.addBreak(BreakType.PAGE);
	
	XWPFParagraph paragraphAlgInfo = document.createParagraph();
	XWPFRun titleAlgInfo = paragraphAlgInfo.createRun();
	titleAlgInfo.setText("1.	ALGEMENE GEGEVENS");
	titleAlgInfo.setFontSize(12);
	titleAlgInfo.setBold(true);
	paragraphAlgInfo.setAlignment(ParagraphAlignment.CENTER);
	paragraphAlgInfo.setSpacingAfter(10);

	XWPFParagraph paragraphInventLijn = document.createParagraph();
	XWPFRun InventLijn = paragraphInventLijn.createRun();
	InventLijn.setText("Inventarisatie Lijn");
	InventLijn.setFontSize(12);
	InventLijn.setBold(true);
	paragraphInventLijn.setAlignment(ParagraphAlignment.LEFT);
	paragraphInventLijn.setSpacingAfter(5);

	XWPFTable table3 = document.createTable();
	XWPFTableRow table3RowOne = table3.getRow(0);
	table3.createRow();
	XWPFTableRow table3RowTwo = table3.getRow(1);
	
	double	textInventLijn = Double.parseDouble(inventLijn.getText() ) * 20 /10000 ;
	
	table3RowOne.getCell(0).setText("Lengte lijn (m)");
	table3RowTwo.getCell(0).setText("Oppervlakte strook (ha)");
	table3RowOne.createCell().setText(inventLijn.getText());
	table3RowTwo.createCell().setText( String.valueOf(textInventLijn));
	
	table3.setCellMargins(50, 400, 50, 400);
	
	setTableAlignment(table3, STJc.CENTER);
	
	XWPFParagraph paragraph2 = document.createParagraph();
	XWPFRun run2 = paragraph2.createRun();
	run2.addBreak();

	XWPFParagraph paragraphOpp = document.createParagraph();
	XWPFRun Opp =paragraphOpp .createRun();
	Opp.setText("Oppervlakte");
	Opp.setFontSize(12);
	Opp.setBold(true);
	paragraphOpp.setAlignment(ParagraphAlignment.LEFT);
	paragraphOpp.setSpacingAfter(5);
	
	XWPFTable table4 = document.createTable();
	XWPFTableRow table4RowOne = table4.getRow(0);
	
	
	table4RowOne.getCell(0).setText("kapvak");
	table4RowOne.createCell().setText("Oppervlakte (ha)");
	
for (int i = 0 ; i < numKvSelect; i++){
	
	table4.createRow();
	XWPFTableRow table4Row = table4.getRow(1+i);
	table4Row.getCell(0).setText(datakvi.get(i));
	table4Row.getCell(1).setText(kvOppTF[i].getText());
		
}
	
table4.setCellMargins(50, 400, 50, 400);
	setTableAlignment(table4, STJc.CENTER);
	
	XWPFParagraph paragraph3 = document.createParagraph();
	XWPFRun run3 = paragraph3.createRun();
	run3.addBreak();

	XWPFParagraph paragraphKvAf = document.createParagraph();
	XWPFRun KvAf = paragraphKvAf.createRun();
	KvAf.setText("Kapvak afbakening");
	KvAf.setFontSize(12);
	KvAf.setBold(true);
	paragraphKvAf.setAlignment(ParagraphAlignment.LEFT);
	paragraphKvAf.setSpacingAfter(5);

	XWPFTable table5 = document.createTable();
	XWPFTableRow table5RowOne = table5.getRow(0);

	table5RowOne.getCell(0).setText("Kapvak");
	table5RowOne.createCell().setText("Kapvak-nummering");
	table5RowOne.createCell().setText("Afbakening lijn");
	for (int i = 0; i < datakvc.size(); i++ ){
	table5.createRow();
	XWPFTableRow table5Row = table5.getRow(1+i);
	table5Row.getCell(0).setText(datakvc.get(i));
	table5Row.getCell(1).setText(datakapvnString.get(i));
	table5Row.getCell(2).setText(dataAfbString.get(i));
	}
	table5.setCellMargins(50, 400, 50, 400);
	setTableAlignment(table5, STJc.CENTER);
	
	XWPFParagraph paragraph4 = document.createParagraph();
	XWPFRun run4 = paragraph4.createRun();
	run4.addBreak(BreakType.PAGE);

	
	XWPFTable table6 = document.createTable();
	XWPFTableRow table6RowOne = table6.getRow(0);

	
	table6RowOne.getCell(0).setText("Kapvak:");
	table6RowOne.createCell().setText("Kapvak markatie:");
	table6RowOne.createCell().setText("kleur:");
	table6RowOne.createCell().setText("Markatie-afstand:");
	for (int i = 0; i < datakvc.size(); i ++ ){
		table6.createRow();
		XWPFTableRow table6Row= table6.getRow(1+i);
		table6Row.getCell(0).setText(datakvc.get(i));
		table6Row.getCell(1).setText(datamarkmethodString.get(i));
		table6Row.getCell(2).setText(datamarkkleur.get(i));
		table6Row.getCell(3).setText(datamark_afst.get(i));
	}
	
	table6.setCellMargins(50, 400, 50, 400);
	setTableAlignment(table6, STJc.CENTER);
	
	XWPFParagraph paragraph5 = document.createParagraph();
	XWPFRun run5 = paragraph5.createRun();
	run5.addBreak();
	
	
	XWPFParagraph paragraphNbosBA = document.createParagraph();
	XWPFRun NbosBA = paragraphNbosBA.createRun();
	NbosBA.setText("Niet bosbouwactiviteiten");
	NbosBA.setFontSize(12);
	NbosBA.setBold(true);
	paragraphNbosBA.setAlignment(ParagraphAlignment.LEFT);
	paragraphNbosBA.setSpacingAfter(5);
	
	XWPFTable table5_5 = document.createTable();
	XWPFTableRow table5_5RowOne = table5_5.getRow(0);
	
	table5_5RowOne.getCell(0).setText("Niet bosbouwactiviteiten");
	table5_5RowOne.createCell().setText(nbosactString);
	
	table5_5.setCellMargins(50, 200, 50, 200);
	setTableAlignment(table5_5, STJc.CENTER);
	
	XWPFParagraph paragraph5_1 = document.createParagraph();
	XWPFRun run5_1 = paragraph5_1.createRun();
	run5_1.addBreak();
	
	if(nbosactString.equals("Ja")){
		XWPFTable tablex = document.createTable();
		XWPFTableRow tablexRowOne = tablex.getRow(0);
		
		tablexRowOne.getCell(0).setText("Kapvak");
		tablexRowOne.createCell().setText("Type activiteiten");
		for (int i = 0; i <datakvc.size(); i++){
		tablex.createRow();
		XWPFTableRow tablexRow = tablex.getRow(1+i);
		tablexRow.getCell(0).setText(datakvc.get(i));
		tablexRow.getCell(1).setText(data_ta.get(i));
		
		}
		tablex.setCellMargins(50, 400, 10, 400);
		setTableAlignment(tablex, STJc.CENTER);
	}
	
	XWPFParagraph paragraph7 = document.createParagraph();
	XWPFRun run7 = paragraph7.createRun();
	run7.addBreak(BreakType.PAGE);
	
	XWPFParagraph paragraphEHKA = document.createParagraph();
	XWPFRun EHKA = paragraphEHKA.createRun();
	EHKA.setText("Eerdere Houtkapactiviteiten");
	EHKA.setFontSize(12);
	EHKA.setBold(true);
	paragraphEHKA.setAlignment(ParagraphAlignment.LEFT);
	paragraphEHKA.setSpacingAfter(5);
	
	XWPFTable table6_5 = document.createTable();
	XWPFTableRow table6_5RowOne = table6_5.getRow(0);
	
	table6_5RowOne.getCell(0).setText("Eerdere houtkapactiviteiten");
	table6_5RowOne.createCell().setText(ehkaString);
	
	table6_5.setCellMargins(50, 200, 50, 200);
	setTableAlignment(table6_5, STJc.CENTER);
	
	XWPFParagraph paragraph5_5 = document.createParagraph();
	XWPFRun run5_5 = paragraph5_5.createRun();
	run5_5.addBreak();
	
	if(ehkaString.equals("Ja")){
	XWPFTable table7 = document.createTable();
	XWPFTableRow table7RowOne = table7.getRow(0);
	
	table7RowOne.getCell(0).setText("Kapvak");
	table7RowOne.createCell().setText("Schattings tijdstip activiteiten");
	table7RowOne.createCell().setText("Schaal activiteiten");
	for (int i = 0; i < datakvc.size(); i++){
	table7.createRow();
	XWPFTableRow table7Row = table7.getRow(1+i);
	table7Row.getCell(0).setText(datakvc.get(i));
	table7Row.getCell(1).setText(datamarkJaarString.get(i));
	table7Row.getCell(2).setText(datamarkSchaalString.get(i));
	}
	table7.setCellMargins(50, 400, 10, 400);
	setTableAlignment(table7, STJc.CENTER);
	}
	
	XWPFParagraph paragraph6 = document.createParagraph();
	XWPFRun run6 = paragraph6.createRun();
	run6.addBreak();
	run6.addBreak();
	
	XWPFParagraph paragraphNStronken = document.createParagraph();
	XWPFRun NStronken = paragraphNStronken.createRun();
	NStronken.setText("Aantal Stronken");
	NStronken.setFontSize(12);
	NStronken.setBold(true);
	paragraphNStronken.setAlignment(ParagraphAlignment.LEFT);
	paragraphNStronken.setSpacingAfter(5);
	
	XWPFTable table8 = document.createTable();
	XWPFTableRow table8RowOne = table8.getRow(0);
	table8.createRow();
	XWPFTableRow table8RowTwo = table8.getRow(1);
	table8.createRow();
	XWPFTableRow table8RowThree = table8.getRow(2);
	
	table8RowOne.getCell(0).setText("Jonger dan 1 jaar");
	table8RowTwo.getCell(0).setText("1 tot 5 jaar");
	table8RowThree.getCell(0).setText("Ouder dan 5 jaar");
	table8RowOne.createCell().setText(String.format("%s", nStronken1)  );
	table8RowTwo.createCell().setText(String.format("%s", nStronken2)  );
	table8RowThree.createCell().setText(String.format("%s", nStronken3)  );
	
	table8.setCellMargins(100, 300, 100, 300);
	setTableAlignment(table8, STJc.CENTER);
	
	XWPFParagraph paragraphOpm = document.createParagraph();
	XWPFRun opmerkingen = paragraphOpm.createRun();
	opmerkingen.setText("Algemene opmerkingen: ");
	opmerkingen.setFontSize(12);
	opmerkingen.setBold(true);
	paragraphOpm.setAlignment(ParagraphAlignment.LEFT);
	paragraphOpm.setSpacingBefore(5);
	paragraphOpm.setSpacingAfter(5);
	
	
	 if (!body.isSetSectPr()) {
	      body.addNewSectPr();
	 }
	 CTSectPr section2 = body.getSectPr();

	 if(!section2.isSetPgSz()) {
	     section2.addNewPgSz();
	 }
	
	XWPFParagraph paragraph21 = document.createParagraph();
	XWPFRun run21 = paragraph21.createRun();
	run21.addBreak(BreakType.PAGE);
	
	
	 String imgFile = imageSel.getText();
		
	  try{
		  
	  if( imgFile.length() > 0 ){

	changeOrientation(document, "PORTRAIT", true);
	 pageSize.setW(BigInteger.valueOf(11900));
	 pageSize.setH(BigInteger.valueOf(16840)); 
	
		
	XWPFParagraph paragraph8 = document.createParagraph();
	XWPFRun run8 = paragraph8.createRun();
	run8.setText("Overzicht Kaart");
	run8.setBold(true);

	 
		  
	XWPFParagraph titleImage = document.createParagraph();    
		XWPFRun runImage = titleImage.createRun();
  
 
			    FileInputStream is = null;
				try {
					is = new FileInputStream(imgFile);
				} catch (FileNotFoundException e1) {
					
					e1.printStackTrace();
				}
			    runImage.addBreak();
			    try {
					runImage.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, imgFile, Units.toEMU(500),Units.toEMU(600));
				
				} catch (InvalidFormatException | IOException e1) {
					
					e1.printStackTrace();
				} 
			    try {
					is.close();
				} catch (IOException e1) {
				
					e1.printStackTrace();
				}
	

	 if (!body.isSetSectPr()) {
	      body.addNewSectPr();
	 }
	 CTSectPr section3 = body.getSectPr();

	 if(!section3.isSetPgSz()) {
	     section3.addNewPgSz();
	 }

		XWPFParagraph paragraph9 = document.createParagraph();
		XWPFRun run9 = paragraph9.createRun();
		run9.addBreak(BreakType.PAGE);
		
		changeOrientation( document, "landscape", true);
	 pageSize.setW(BigInteger.valueOf(16840));
	 pageSize.setH(BigInteger.valueOf(11900));
	  } else { 	 stringDisplay0.setText("Geen Overzichtskaart gevonden!");
	   p.add(stringDisplay0);}
	  
	  }catch(Exception img){ JOptionPane.showMessageDialog(null, "Error:2804 (OverzichtsKaart) Fout opgetreden: " + img.getMessage() );}
	  
	  
	XWPFParagraph paragraphTerKen = document.createParagraph();
	XWPFRun TerKen = paragraphTerKen.createRun();
	TerKen.setText("2.	TERREINKENMERKEN");
	TerKen.setFontSize(12);
	TerKen.setBold(true);
	paragraphTerKen.setAlignment(ParagraphAlignment.CENTER);
	paragraphTerKen.setSpacingAfter(10);
	
	XWPFTable table9 = document.createTable();
	XWPFTableRow table9RowOne = table9.getRow(0);
	table9.createRow();
	XWPFTableRow table9RowTwo = table9.getRow(1);
	
	table9RowOne.getCell(0);
	table9RowOne.createCell().setText("Bostype");
	table9RowOne.addNewTableCell();
	table9RowOne.createCell().setText("Ondergroei");
	table9RowOne.createCell().setText("Bodemtype");
	table9RowOne.addNewTableCell();
	table9RowOne.createCell().setText("Topografie");
	table9RowOne.addNewTableCell();
	table9RowTwo.getCell(0).setText("Kapvak");
	table9RowTwo.createCell().setText("Inventarisatie");
	table9RowTwo.createCell().setText("Controle");
	table9RowTwo.createCell().setText("Controle");
	table9RowTwo.createCell().setText("Inventarisatie");
	table9RowTwo.createCell().setText("Controle");
	table9RowTwo.createCell().setText("Inventarisatie");
	table9RowTwo.createCell().setText("Controle");

	mergeCellHorizontally(table9, 0, 1,2);
	mergeCellHorizontally(table9, 0, 4,5);
	mergeCellHorizontally(table9, 0, 6,7);
	
	for(int i = 0; i < datakvbtype.size(); i ++  ){
		table9.createRow();
		XWPFTableRow table9Row = table9.getRow(2 + i);
		table9Row.getCell(0).setText(datakvc.get(i));
		table9Row.getCell(1).setText(databosi.get(i));
		table9Row.getCell(2).setText(databosc.get(i));
		table9Row.getCell(3).setText(dataondergroei.get(i));
		table9Row.getCell(4).setText(databodemi.get(i));
		table9Row.getCell(5).setText(databodemc.get(i));
		table9Row.getCell(6).setText(datatopoi.get(i));
		table9Row.getCell(7).setText(datatopoc.get(i));
		
	}
		table9.setCellMargins(100, 100,100,600);
		setTableAlignment(table9, STJc.CENTER);
	
		XWPFParagraph paragraph10 = document.createParagraph();
		XWPFRun run10 = paragraph10.createRun();
		run10.addBreak(BreakType.PAGE);

		XWPFParagraph paragraphInvent = document.createParagraph();
		XWPFRun Invent = paragraphInvent.createRun();
		Invent.setText("3.	INVENTARISATIE GEGEVENS");
		Invent.setFontSize(12);
		Invent.setBold(true);
		paragraphInvent.setAlignment(ParagraphAlignment.CENTER);
		paragraphInvent.setSpacingAfter(10);
	
	XWPFParagraph paragraphBoomNum = document.createParagraph();
	XWPFRun BoomNum = paragraphBoomNum.createRun();
	BoomNum.setText("Boomnummering");
	BoomNum.setFontSize(12);
	BoomNum.setBold(true);
	paragraphBoomNum.setAlignment(ParagraphAlignment.LEFT);
	paragraphBoomNum.setSpacingAfter(5);
	
	XWPFTable table10 = document.createTable();
	XWPFTableRow table10RowOne = table10.getRow(0);
	table10.createRow();
	XWPFTableRow table10RowTwo = table10.getRow(1);
	table10.createRow();
	XWPFTableRow table10RowThree = table10.getRow(2);
	
	table10RowOne.getCell(0).setText("Boomnummering");
	table10RowOne.createCell().setText("Aantal bomen");
	table10RowOne.createCell().setText("%");
	table10RowTwo.getCell(0).setText("Aanwezig");
	table10RowTwo.createCell().setText(String.valueOf(boomValid.size()));
	table10RowTwo.createCell().setText(String.valueOf(Math.round(procentboomValid)));
	table10RowThree.getCell(0).setText("Niet aanwezig");
	table10RowThree.createCell().setText(String.valueOf( boomna.size()));
	table10RowThree.createCell().setText(String.valueOf(Math.round(procentboomna)));
	
	XWPFParagraph paragraphToel1 = document.createParagraph();
	XWPFRun Toel1 = paragraphToel1.createRun();
	Toel1.setText("Toelichting:");
	Toel1.setFontSize(12);
	Toel1.setBold(true);
	paragraphToel1.setAlignment(ParagraphAlignment.LEFT);
	paragraphToel1.setSpacingAfter(5);
	
	table10.setCellMargins(100, 100, 100,600);
	setTableAlignment(table10, STJc.CENTER);
	
	XWPFParagraph paragraph11 = document.createParagraph();
	XWPFRun run11 = paragraph11.createRun();
	run11.addBreak();

	XWPFParagraph paragraphHS = document.createParagraph();
	XWPFRun HS = paragraphHS.createRun();
	HS.setText("Houtsoort");
	HS.setFontSize(12);
	HS.setBold(true);
	paragraphHS.setAlignment(ParagraphAlignment.LEFT);
	paragraphHS.setSpacingAfter(5);
	
	XWPFTable table11 = document.createTable();
	XWPFTableRow table11RowOne = table11.getRow(0);
	table11.createRow();
	XWPFTableRow table11RowTwo = table11.getRow(1);
	table11.createRow();
	XWPFTableRow table11RowThree = table11.getRow(2);

	table11RowOne.getCell(0).setText("Identificatie");
	table11RowOne.createCell().setText("Aantal bomen");
	table11RowOne.createCell().setText("%");
	table11RowTwo.getCell(0).setText("Juist");
	table11RowTwo.createCell().setText(String.valueOf(countboomcodes));
	table11RowTwo.createCell().setText(String.valueOf(Math.round(countboomcodesproc)));
	table11RowThree.getCell(0).setText("Onjuist");
	table11RowThree.createCell().setText(String.valueOf(foutboom));
	table11RowThree.createCell().setText(String.valueOf(Math.round(foutboomproc)));

	table11.setCellMargins(100, 100, 100,200);
	setTableAlignment(table11, STJc.CENTER);
	
	XWPFParagraph paragraphToel2 = document.createParagraph();
	XWPFRun Toel2 = paragraphToel2.createRun();
	Toel2.setText("Toelichting:");
	Toel2.setFontSize(12);
	Toel2.setBold(true);
	if(nconbomen > 0){
	XWPFParagraph paragraphToel2_note = document.createParagraph();
	XWPFRun Toel2_note = paragraphToel2_note.createRun();
	Toel2_note.setText(String.format("•   %s %s %s", "Er zijn", nconbomen, "bomen van de inventarisatiecontrole die geen link maken met de inventarisatiedata"));
	Toel2_note.setFontSize(12);
	paragraphToel2_note.setAlignment(ParagraphAlignment.LEFT);
	paragraphToel2_note.setSpacingAfter(5);
	} else{
	paragraphToel2.setAlignment(ParagraphAlignment.LEFT);
	paragraphToel2.setSpacingAfter(5);}
	
	
	XWPFParagraph paragraph12 = document.createParagraph();
	XWPFRun run12 = paragraph12.createRun();
	run12.addBreak(BreakType.PAGE);
	
	XWPFParagraph paragraphDuurzaam = document.createParagraph();
	XWPFRun Duurzaam = paragraphDuurzaam.createRun();
	Duurzaam.setText("Inachtneming duurzaamheidregel");
	Duurzaam.setFontSize(12);
	Duurzaam.setBold(true);
	paragraphDuurzaam.setAlignment(ParagraphAlignment.LEFT);
	paragraphDuurzaam.setSpacingAfter(5);
	
	XWPFTable table12 = document.createTable();
	XWPFTableRow table12RowOne = table12.getRow(0);
	table12.createRow();
	XWPFTableRow table12RowTwo = table12.getRow(1);
	table12.createRow();
	XWPFTableRow table12RowThree = table12.getRow(2);
	table12.createRow();
	XWPFTableRow table12RowFour = table12.getRow(3);
	table12.createRow();
	XWPFTableRow table12RowFive = table12.getRow(4);
	table12.createRow();
	XWPFTableRow table12RowSix = table12.getRow(5);
	table12.createRow();
	XWPFTableRow table12RowSeven = table12.getRow(6);
	
	table12RowOne.getCell(0).setText("Categorie");
	table12RowOne.createCell().setText("Aantal bomen");
	table12RowTwo.getCell(0).setText("0: niet gemarkeerd");
	table12RowTwo.createCell().setText(String.valueOf(mark0));
	table12RowThree.getCell(0).setText("1: alles ok");
	table12RowThree.createCell().setText(String.valueOf(mark1));
	table12RowFour.getCell(0).setText("2: foutieve 10m regel");
	table12RowFour.createCell().setText(String.valueOf(mark2));
	table12RowFive.getCell(0).setText("3: dichter dan 20m bij een kreek");
	table12RowFive.createCell().setText(String.valueOf(mark3));
	table12RowSix.getCell(0).setText("4: dichter dan 20m bij een zwamp");
	table12RowSix.createCell().setText(String.valueOf(mark4));
	table12RowSeven.getCell(0).setText("5: helling steiler dan 30% (17 graden)");
	table12RowSeven.createCell().setText(String.valueOf(mark5));
	
	table12.setCellMargins(100, 100, 100,200);
	setTableAlignment(table12, STJc.CENTER);
	
	XWPFParagraph paragraph13 = document.createParagraph();
	XWPFRun run13= paragraph13.createRun();
	run13.addBreak();
	
	XWPFParagraph paragraphDBH = document.createParagraph();
	XWPFRun DBH = paragraphDBH.createRun();
	DBH.setText("DBH");
	DBH.setFontSize(12);
	DBH.setBold(true);
	paragraphDBH.setAlignment(ParagraphAlignment.LEFT);
	paragraphDBH.setSpacingAfter(5);
	
	XWPFTable table13 = document.createTable();
	XWPFTableRow table13RowOne = table13.getRow(0);
	table13.createRow();
	XWPFTableRow table13RowTwo = table13.getRow(1);
	table13.createRow();
	XWPFTableRow table13RowThree = table13.getRow(2);
	table13.createRow();
	XWPFTableRow table13RowFour = table13.getRow(3);

	table13RowOne.getCell(0).setText("Klasse");
	table13RowOne.createCell().setText("Aantal bomen");
	table13RowOne.createCell().setText("in %");
	table13RowTwo.getCell(0).setText("<= 5 cm");
	table13RowTwo.createCell().setText(String.valueOf(sizeDbhClass1));
	table13RowTwo.createCell().setText(String.valueOf(Math.round(sizeDbhClass1_prc)));
	table13RowThree.getCell(0).setText("5cm t/m 10 cm");
	table13RowThree.createCell().setText(String.valueOf(sizeDbhClass2));
	table13RowThree.createCell().setText(String.valueOf(Math.round(sizeDbhClass2_prc)));
	table13RowFour.getCell(0).setText("> 10 cm");
	table13RowFour.createCell().setText(String.valueOf(sizeDbhClass3));
	table13RowFour.createCell().setText(String.valueOf(Math.round(sizeDbhClass3_prc)));

	table13.setCellMargins(100, 100, 100,200);
	setTableAlignment(table13, STJc.CENTER);
	
	XWPFParagraph paragraphToel3 = document.createParagraph();
	XWPFRun Toel3 = paragraphToel3.createRun();
	Toel3.setText("Toelichting:");
	Toel3.setFontSize(12);
	Toel3.setBold(true);
	paragraphToel3.setAlignment(ParagraphAlignment.LEFT);
	paragraphToel3.setSpacingAfter(5);
	
	
	XWPFParagraph paragraphHoogteCom = document.createParagraph();
	XWPFRun HoogteCom = paragraphHoogteCom.createRun();
	HoogteCom.setText("Hoogte");
	HoogteCom.setFontSize(12);
	HoogteCom.setBold(true);
	paragraphHoogteCom.setAlignment(ParagraphAlignment.LEFT);
	paragraphHoogteCom.setSpacingAfter(5);
	
	XWPFTable table14 = document.createTable();
	XWPFTableRow table14RowOne = table14.getRow(0);
	table14.createRow();
	XWPFTableRow table14RowTwo = table14.getRow(1);
	table14.createRow();
	XWPFTableRow table14RowThree = table14.getRow(2);
	table14.createRow();
	XWPFTableRow table14RowFour = table14.getRow(3);
	
	table14RowOne.getCell(0).setText("Klasse");
	table14RowOne.createCell().setText("Aantal bomen");
	table14RowOne.createCell().setText("In %");
	table14RowTwo.getCell(0).setText(" <= 2m");
	table14RowTwo.createCell().setText(String.valueOf(sizeHcClass1));
	table14RowTwo.createCell().setText(String.valueOf(Math.round(sizeHcClass1_prc)));
	table14RowThree.getCell(0).setText("2m t/m 5m");
	table14RowThree.createCell().setText(String.valueOf(sizeHcClass2));
	table14RowThree.createCell().setText(String.valueOf(Math.round(sizeHcClass2_prc)));
	table14RowFour.getCell(0).setText(">  5m");
	table14RowFour.createCell().setText(String.valueOf(sizeHcClass3));
	table14RowFour.createCell().setText(String.valueOf(Math.round(sizeHcClass3_prc)));

	table14.setCellMargins(100, 100, 100,200);
	setTableAlignment(table14, STJc.CENTER);
	
	XWPFParagraph paragraphToel4 = document.createParagraph();
	XWPFRun Toel4 = paragraphToel4.createRun();
	Toel4.setText("Toelichting:");
	Toel4.setFontSize(12);
	Toel4.setBold(true);
	paragraphToel4.setAlignment(ParagraphAlignment.LEFT);
	paragraphToel4.setSpacingAfter(5);
		
	XWPFParagraph paragraph15 = document.createParagraph();
	XWPFRun run15 = paragraph15.createRun();
	run15.addBreak();

	XWPFParagraph paragraphVolume = document.createParagraph();
	XWPFRun Volume = paragraphVolume.createRun();
	 Volume.setText("Volume");
	 Volume.setFontSize(12);
	 Volume.setBold(true);
	paragraphVolume.setAlignment(ParagraphAlignment.LEFT);
	paragraphVolume.setSpacingAfter(5);
	
	
	XWPFTable table15 = document.createTable();
	XWPFTableRow table15RowOne = table15.getRow(0);


	table15RowOne.getCell(0).setText("Kapvak");
	table15RowOne.createCell().setText("Geselecteerd (m3/ha)" );
	table15RowOne.createCell().setText("niet geselecteerd (m3/ha)" );
	table15RowOne.createCell().setText("Verboden (m3/ha)" );
	
	if (datakvi.isEmpty()) {
		table15.createRow();
		XWPFTableRow table15Row = table15.getRow(1);
	table15Row.getCell(0).setText("0");
	table15Row.getCell(1).setText("0");
	table15Row.getCell(2).setText("0");
	table15Row.getCell(3).setText("0");
	} else {
		
	for(int i = 0; i < numKvSelect; i ++  ){
		table15.createRow();
		XWPFTableRow table15Row = table15.getRow(1+i);
	table15Row.getCell(0).setText(datakvi.get(i));
	try{ table15Row.getCell(1).setText( String.valueOf(Math.round(volArraySel1.get(i)*100.0)/100.0) );    }catch(Exception vol1){table15Row.getCell(1).setText("0");}
	
	try{ table15Row.getCell(2).setText( String.valueOf(Math.round(volArraySel2.get(i)*100.0)/100.0) );    }catch(Exception vol2){table15Row.getCell(2).setText("0");}
	
	try{ table15Row.getCell(3).setText( String.valueOf(Math.round(volArraySel3.get(i)*100.0)/100.0) );	  }catch(Exception vol3){table15Row.getCell(3).setText("0");}
	}
}
	table15.setCellMargins(100, 100, 100,200);
	setTableAlignment(table15, STJc.CENTER);
	
	XWPFParagraph paragraphToel5 = document.createParagraph();
	XWPFRun Toel5 = paragraphToel5.createRun();
	Toel5.setText("Toelichting:");
	Toel5.setFontSize(12);
	Toel5.setBold(true);
	paragraphToel5.setAlignment(ParagraphAlignment.LEFT);
	paragraphToel5.setSpacingAfter(5);
	
	XWPFParagraph paragraph16 = document.createParagraph();
	XWPFRun run16 = paragraph16.createRun();
	run16.addBreak(BreakType.PAGE);
	
	XWPFParagraph paragraphVerboden = document.createParagraph();
	XWPFRun Verboden = paragraphVerboden.createRun();
	Verboden.setText("Selectie verboden soorten");
	Verboden.setFontSize(12);
	Verboden.setBold(true);
	paragraphVerboden.setAlignment(ParagraphAlignment.LEFT);
	paragraphVerboden.setSpacingAfter(5);
	
	XWPFTable table16 = document.createTable();
	XWPFTableRow table16RowOne = table16.getRow(0);

	table16RowOne.getCell(0).setText("Kapvak");
	table16RowOne.createCell().setText("Houtsoort(en)");
	table16RowOne.createCell().setText("Aantal bomen");
	if(dataKv_verb.isEmpty()){
		table16.createRow();
		XWPFTableRow table16Row = table16.getRow(1);
		table16Row.getCell(0).setText("0");
		table16Row.getCell(1).setText("0");
		table16Row.getCell(2).setText("0");
	} else{
	for(int i = 0; i < dataKv_verb.size(); i ++  ){
	table16.createRow();
	XWPFTableRow table16Row = table16.getRow(1+i);
	table16Row.getCell(0).setText(dataKv_verb.get(i).toString());
	table16Row.getCell(1).setText(dataHs_verb.get(i).toString());
	table16Row.getCell(2).setText(dataNum_verb.get(i).toString());
	}
	}
	
	table16.setCellMargins(100, 100, 100,200);
	setTableAlignment(table16, STJc.CENTER);
	
	XWPFParagraph paragraphAdvies1 = document.createParagraph();
	XWPFRun Advies1 = paragraphAdvies1.createRun();
	Advies1.setText("Advies: Check als er inderdaad ontheffing is aangevraagd/verleend");
	Advies1.setFontSize(12);
	Advies1.setBold(true);
	paragraphAdvies1.setAlignment(ParagraphAlignment.LEFT);
	paragraphAdvies1.setSpacingAfter(5);
	
	XWPFParagraph paragraph17 = document.createParagraph();
	XWPFRun run17 = paragraph17.createRun();
	run17.addBreak(BreakType.PAGE);
	
	XWPFParagraph paragraphOndermaats = document.createParagraph();
	XWPFRun Ondermaats = paragraphOndermaats.createRun();
	Ondermaats.setText("Selectie ondermaatse bomen");
	Ondermaats.setFontSize(12);
	Ondermaats.setBold(true);
	paragraphOndermaats.setAlignment(ParagraphAlignment.LEFT);
	paragraphOndermaats.setSpacingAfter(5);
	
	XWPFTable table17 = document.createTable();
	XWPFTableRow table17RowOne = table17.getRow(0);
	
	table17RowOne.getCell(0).setText("Kapvak");
	table17RowOne.createCell().setText("Houtsoort(en)");
	table17RowOne.createCell().setText("Aantal bomen");
	if (dataKv_onderma.isEmpty()){
		table17.createRow();
		XWPFTableRow table17Row = table17.getRow(1);
		table17Row.getCell(0).setText("0");
		table17Row.getCell(1).setText("0");
		table17Row.getCell(2).setText("0");
	} else {
	for(int i = 0; i < dataKv_onderma.size(); i ++  ){
	table17.createRow();
	XWPFTableRow table17Row = table17.getRow(1+i);
	table17Row.getCell(0).setText(dataKv_onderma.get(i).toString());
	table17Row.getCell(1).setText(dataHs_onderma.get(i).toString());
	table17Row.getCell(2).setText(dataNum_onderma.get(i).toString());
	}
}	
	table17.setCellMargins(100, 100, 100,200);
	setTableAlignment(table17, STJc.CENTER);
	
	XWPFParagraph paragraphAdvies2 = document.createParagraph();
	XWPFRun Advies2 = paragraphAdvies2.createRun();
	Advies2.setText("Advies: Check als er inderdaad ontheffing is aangevraagd/verleend");
	Advies2.setFontSize(12);
	Advies2.setBold(true);
	paragraphAdvies2.setAlignment(ParagraphAlignment.LEFT);
	paragraphAdvies2.setSpacingAfter(5);
	
	
	XWPFParagraph paragraph18 = document.createParagraph();
	XWPFRun run18 = paragraph18.createRun();
	run18.addBreak(BreakType.PAGE);

	XWPFParagraph paragraphSelvsMar = document.createParagraph();
	XWPFRun SelvsMar = paragraphSelvsMar.createRun();
	SelvsMar.setText("Selectie vs Markatie");
	SelvsMar.setFontSize(12);
	SelvsMar.setBold(true);
	paragraphSelvsMar.setAlignment(ParagraphAlignment.LEFT);
	paragraphSelvsMar.setSpacingAfter(5);
	
	XWPFTable table18 = document.createTable();
	XWPFTableRow table18RowOne = table18.getRow(0);
	table18.createRow();
	XWPFTableRow table18RowTwo = table18.getRow(1);
	
	
	table18RowOne.getCell(0);
	table18RowOne.createCell().setText("Wel gemarkeerd & wel geselecteerd");
	table18RowOne.createCell().setText("Niet gemarkeerd & wel geselecteerd");
	table18RowOne.createCell().setText("Wel gemarkeerd & niet geselecteerd");
	table18RowOne.createCell().setText("niet gemarkeerd & niet geselecteerd");
	table18RowTwo.getCell(0).setText("Kapvak");
	table18RowTwo.createCell().setText("Aantal bomen");
	table18RowTwo.createCell().setText("Aantal bomen");
	table18RowTwo.createCell().setText("Aantal bomen");
	table18RowTwo.createCell().setText("Aantal bomen");
	
	if(markselKV.isEmpty() ){
	
		table18.createRow();
		XWPFTableRow table18Row = table18.getRow(2);
		table18Row.getCell(0).setText("0");
		table18Row.getCell(1).setText("0");
		table18Row.getCell(2).setText("0");
		table18Row.getCell(3).setText("0");
		table18Row.getCell(4).setText("0");
		
	} else { 
	for(int i = 0; i < markselKV.size(); i ++  ){

		table18.createRow();
		XWPFTableRow table18Row = table18.getRow(2+i);
		table18Row.getCell(0).setText(markselKV.get(i).toString());
		table18Row.getCell(1).setText(mark1sel1.get(i).toString());
		table18Row.getCell(2).setText(mark0sel1.get(i).toString());
		table18Row.getCell(3).setText(mark1sel0.get(i).toString());
		table18Row.getCell(4).setText(mark0sel0.get(i).toString());
	}
}
	
	table18.setCellMargins(100, 100, 100,200);
	setTableAlignment(table18, STJc.CENTER);
	
	XWPFParagraph paragraph19 = document.createParagraph();
	XWPFRun run19 = paragraph19.createRun();
	run19.addBreak();
	
	XWPFTable table19 = document.createTable();
	XWPFTableRow table19RowOne = table19.getRow(0);
	
	table19RowOne.getCell(0).setText("Kapvak");
	table19RowOne.createCell().setText("Markatie waargenomen in kapvak");
	table19RowOne.createCell().setText("Sleepwegen markatie waargenomen in kapvak");
	table19RowOne.createCell().setText("Opmerkingen");
	
	
	for(int i = 0; i < datasleepmarkString.size() ; i ++){
		table19.createRow();
		XWPFTableRow table19Row = table19.getRow(1+i);
	
	table19Row.getCell(0).setText(datakvc.get(i));
	table19Row.getCell(1).setText(databoommarkString.get(i));
	table19Row.getCell(2).setText(datasleepmarkString.get(i));
	table19Row.getCell(3);
	}
	
	table19.setCellMargins(100, 100, 100,200);
	setTableAlignment(table19, STJc.CENTER);
	
	XWPFParagraph paragraph20 = document.createParagraph();
	XWPFRun run20 = paragraph20.createRun();
	run20.addBreak(BreakType.PAGE);
	
	XWPFParagraph paragraphConclusie = document.createParagraph();
	XWPFRun ConclusieRun = paragraphConclusie.createRun();
	ConclusieRun.setText("CONCLUSIE");
	ConclusieRun.setFontSize(14);
	ConclusieRun.setBold(true);
	paragraphConclusie.setAlignment(ParagraphAlignment.CENTER);
	paragraphConclusie.setBorderBottom(Borders.THICK);	
	
	
	
	try {
		document.write(out);
	} catch (IOException e) {
	
		JOptionPane.showMessageDialog(null, "Error: 3063 Fout opgetreden: " +  e.getMessage() );
	}
    try {
		document.close();
	} catch (IOException e) {
		
		e.printStackTrace();
	}
		try {
			out.close();
		} catch (IOException e) {
		
			e.printStackTrace();
		}
	
	

		
}catch (Exception rap){	 //JOptionPane.showMessageDialog(null, "Error: Rapport schrijven Fout opgetreden: " +  rap.getStackTrace() );
rap.printStackTrace();
}
JOptionPane.showMessageDialog(null, "Rapport Voltooid!");
	//		   } catch(Exception err){JOptionPane.showMessageDialog(null, "Error: Fatal error" +  err.getMessage() ); }
			   
f.setAlwaysOnTop(false);
f.setVisible(false);

		   }
			   
	
		}
	   
	   static void mergeCellHorizontally(XWPFTable table, int row, int fromCol, int toCol) {
		   for(int colIndex = fromCol; colIndex <= toCol; colIndex++){
		    CTHMerge hmerge = CTHMerge.Factory.newInstance();
		    if(colIndex == fromCol){
		     // The first merged cell is set with RESTART merge value
		     hmerge.setVal(STMerge.RESTART);
		    } else {
		     // Cells which join (merge) the first one, are set with CONTINUE
		     hmerge.setVal(STMerge.CONTINUE);
		    }
		    XWPFTableCell cell = table.getRow(row).getCell(colIndex);
		    // Try getting the TcPr. Not simply setting an new one every time.
		    CTTcPr tcPr = cell.getCTTc().getTcPr();
		    if (tcPr != null) {
		     tcPr.setHMerge(hmerge);
		    } else {
		     // only set an new TcPr if there is not one already
		     tcPr = CTTcPr.Factory.newInstance();
		     tcPr.setHMerge(hmerge);
		     cell.getCTTc().setTcPr(tcPr);
		    }
		   }
		  }
	   
		 public static void setTableAlignment(XWPFTable table, STJc.Enum justification) {
			    CTTblPr tblPr = table.getCTTbl().getTblPr();
			    CTJc jc = (tblPr.isSetJc() ? tblPr.getJc() : tblPr.addNewJc());
			    jc.setVal(justification);
			}
		 
		 private static void changeOrientation(XWPFDocument document, String orientation, boolean pFinalSection){
			    CTSectPr section;
			    if (pFinalSection) {
			        CTDocument1 doc = document.getDocument();
			        CTBody body = doc.getBody();
			        section = body.getSectPr() != null ? body.getSectPr() : body.addNewSectPr();
			        XWPFParagraph para = document.createParagraph();
			        CTP ctp = para.getCTP();
			        CTPPr br = ctp.addNewPPr();
			        br.setSectPr(section);
			    } else {
			        XWPFParagraph para = document.createParagraph();
			        CTP ctp = para.getCTP();
			        CTPPr br = ctp.addNewPPr();
			        section = br.addNewSectPr();
			        br.setSectPr(section);
			    }
			    CTPageSz pageSize = section.isSetPgSz() ? section.getPgSz() : section.addNewPgSz();
			    if(orientation.equals("landscape")){
			        pageSize.setOrient(STPageOrientation.LANDSCAPE);
			        pageSize.setW(BigInteger.valueOf(842 * 20));
			        pageSize.setH(BigInteger.valueOf(595 * 20));
			    }
			    else{
			        pageSize.setOrient(STPageOrientation.PORTRAIT);
			        pageSize.setH(BigInteger.valueOf(842 * 20));
			        pageSize.setW(BigInteger.valueOf(595 * 20));
			    }
			}
		 		  		 
	} //Gui class closing bracket
	  

