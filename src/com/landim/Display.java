package com.landim;

import java.awt.event.*;
import javax.swing.*;
import java.awt.Dimension;
import java.io.*;


public class Display extends JPanel implements ActionListener{

	private JFileChooser fileChooser;
	private JButton openButton; 
	private BufferedReader br;
	private File file;
	int returnVal;
	
	public Display() {
		fileChooser = new JFileChooser();
		openButton = new JButton("Select");
		
		setPreferredSize(new Dimension(278, 179));
		setLayout(null);
		
		add(openButton); 
		
		openButton.setBounds(84, 145, 100, 25);
		openButton.addActionListener(this);
	}
	
	
	@Override
	public void actionPerformed(ActionEvent e) {
		if(e.getSource() == openButton){
			returnVal = fileChooser.showOpenDialog(null);
			if (returnVal == JFileChooser.APPROVE_OPTION){
				file = fileChooser.getSelectedFile();	
			}
		}	
	}

}
