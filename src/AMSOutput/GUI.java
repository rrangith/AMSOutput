package AMSOutput;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;


import java.io.*;
import java.util.ArrayList;

import java.util.LinkedList;


/**
 * Created by rahul on 2017-07-12.
 */

public class GUI extends JFrame {

    /**Text Fields**/
    private JTextField fileOneField; //name of Bill of Material File
    private JTextField fileTwoField; //name of X Y Coordinates File
    private JTextField fileOnePartField; //starting part number cell
    private JTextField fileOneDesField; //starting designator cell
    private JTextField fileOneLastRowField; //last row to avoid footers/unneeded lines

    private JPanel panel;
    private JButton helpButton;
    private JTextArea area;

    /**File Names**/
    String fileOneString = "";
    String fileTwoString = "";

    File directory; //used to go back to old spot in file chooser

    GUI() {
        super("AMSOutput"); //title of frame
        this.setSize(800, 300); //size: 800px wide, 300px high
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setResizable(true); //makes frame resizable

        panel = new JPanel(); //main panel to add everything to
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS)); //makes everything go from top to bottom

        JLabel instructionsLabel = new JLabel("Select Two Files. "); //instructions

        JPanel inputPanel = new JPanel(); //panel used to get inputs
        inputPanel.setLayout(new BoxLayout(inputPanel, BoxLayout.Y_AXIS)); //makes everything go from top to bottomJOptionPane.showMessageDialog(null, "Error");

        JPanel fileOnePanel = new JPanel(); //panel to get first file's name
        fileOnePanel.setLayout(new FlowLayout()); //makes everything go from left to right
        JLabel fileOneLabel = new JLabel("BOM File (Excel)"); //instructions
        fileOneField = new JTextField(40); //makes text field big enough to display 40 characters
        fileOnePanel.add(fileOneLabel); //adds label
        fileOnePanel.add(fileOneField); //adds text field
        JButton fileOneBut = new JButton("Select"); //adds button to choose file
        fileOneBut.addActionListener(new FileOneButListener()); //adds listener to open file chooser gui
        fileOnePanel.add(fileOneBut); //adds button

        JPanel fileOneInputPanel = new JPanel(); //panel to get starting cells and final row
        fileOneInputPanel.setLayout(new FlowLayout()); //makes everything go from left to right

        JLabel fileOneStart = new JLabel("Starting Part Number Cell:"); //instructions
        fileOnePartField = new JTextField(5); //text field enough to display 5 characters
        fileOneInputPanel.add(fileOneStart); //adds label
        fileOneInputPanel.add(fileOnePartField); //adds text field

        JLabel fileOneStartTwo = new JLabel("Starting Designator Cell"); //instructions
        fileOneDesField = new JTextField(5); //text field enough to display 5 characters
        fileOneInputPanel.add(fileOneStartTwo); //adds label
        fileOneInputPanel.add(fileOneDesField); //adds text field

        JLabel fileOneLastRow = new JLabel("Last Row Number"); //instructions
        fileOneLastRowField = new JTextField(5); //text field enough to display 5 characters
        fileOneInputPanel.add(fileOneLastRow); //adds label
        fileOneInputPanel.add(fileOneLastRowField);//adds text field


        JPanel fileTwoPanel = new JPanel(); //panel to get second file's name
        fileTwoPanel.setLayout(new FlowLayout()); //makes everything go from left to right

        JLabel fileTwoLabel = new JLabel("XY Cor File (CSV)"); //instructions
        fileTwoField = new JTextField(40); //text field enough to display 40 characters
        JButton fileTwoBut = new JButton("Select"); //button to select second file
        fileTwoBut.addActionListener(new FileTwoButListener()); //used to open file chooser gui
        fileTwoPanel.add(fileTwoLabel); //adds label
        fileTwoPanel.add(fileTwoField); //adds field
        fileTwoPanel.add(fileTwoBut); //adds button

        inputPanel.add(fileOnePanel); //adds panel which contains first file's name
        inputPanel.add(fileOneInputPanel); //adds panel which contains starting cells and final row
        inputPanel.add(fileTwoPanel); //adds panel which contains second file's name

        JPanel startPanel = new JPanel(new FlowLayout()); //makes panel to hold start button

        JButton startButton = new JButton("Start"); //makes button
        startButton.addActionListener(new StartButtonListener()); //adds function to the button

        helpButton = new JButton("Help");
        helpButton.addActionListener(new HelpButtonListener());

        startPanel.add(startButton); //adds button to panel
        startPanel.add(helpButton);

        panel.add(instructionsLabel); //adds label to main panel
        panel.add(inputPanel); //adds all input to main panel
        panel.add(startPanel); //adds start button to main panel

        this.add(panel); //adds main panel to frame
        this.setVisible(true); //makes the fram visible
    }

    public class StartButtonListener implements ActionListener { //inner class, used for the functionality of the start button

        /**Variables from frame that will be used**/
        String fieldOne; //first file's name
        String fieldTwo; //second file's name
        String partStart; //starting part number cell
        String desStart; //starting designator cell
        int lastRow; //last row that will be used

        @Override
        public void actionPerformed(ActionEvent e) {

            /**Get file names**/
            fieldOne = fileOneField.getText(); //file 1 name
            fieldTwo = fileTwoField.getText(); //file 2 name
            partStart = fileOnePartField.getText().toUpperCase(); //starting part number cell
            desStart = fileOneDesField.getText().toUpperCase(); //starting designator cell
            lastRow = Integer.parseInt(fileOneLastRowField.getText()); //last row of excel file


            //error check
            if (fieldOne.length() > 4 && fieldTwo.length() > 4) { //makes sure file names are long enough
                if ((fieldOne.substring(fieldOne.length() - 5).equals(".xlsx") || (fieldOne.substring(fieldOne.length() - 4).equalsIgnoreCase(".xls")) && fieldTwo.substring(fieldTwo.length() - 4).equals(".csv"))) { //makes sure extensions are correct
                    try {
                        //make files
                        File file = new File(fieldOne);
                        File fileTwo = new File(fieldTwo);


                        InputStream fs = new FileInputStream(file); //makes input stream
                        Workbook wb = WorkbookFactory.create(fs); //creates workbook
                        Sheet sheet = wb.getSheetAt(0); //gets the first sheet

                        File outputFile = new File(fileTwo.getParentFile().getAbsolutePath() + File.separator + "Differences.txt"); //makes file to output differences
                        FileOutputStream outputFileStream = new FileOutputStream(outputFile); //makes output stream

                        PrintWriter output = new PrintWriter(outputFileStream); //print writer to write to file

                        Row row; //declares row variable

                        /**Declares and initializes cells**/
                        Cell cell; //cell to be compared to other data
                        Cell startingPartCell = null; //starting part number cell
                        Cell startDesCell = null; //starting designator cell
                        Cell partCell = null; //part number cell to be compared to other data
                        Cell desCel = null; //designator cell to be compared to other data

                        int rows = lastRow; //sets the amount of rows based on user input

                        int cols = 0; // No of columns
                        int tmp = 0; //used to get columns

                        /*Variables to split starting cells into letters and numbers for error checking*/
                        String partStartLet = ""; //letters of starting part number cell
                        String partStartNum = ""; //numbers of starting part number cell
                        String desStartLet = ""; //letters of starting designator cell
                        String desStartNum = ""; //numbers of starting designator cell

                        int timesFound = 0; //used to count number of times each designator is found

                        ArrayList<String> updated = new ArrayList<String>(); //used to store lines from file when designators are the same
                        ArrayList<String> newList = new ArrayList<String>(); //used to store lines that will be written to updated output file

                        /**Splits the starting part number cell into numbers and letters**/
                        for (int i = 0; i < partStart.length(); i++) {
                            if (Character.isLetter(partStart.charAt(i))) {
                                partStartLet += partStart.charAt(i);
                            } else if (Character.isDigit(partStart.charAt(i))) {
                                partStartNum += partStart.charAt(i);
                            }
                        }

                        /**Splits the starting designator cell into numbers and letters**/
                        for (int i = 0; i < desStart.length(); i++) {
                            if (Character.isLetter(desStart.charAt(i))) {
                                desStartLet += desStart.charAt(i);
                            } else if (Character.isDigit(desStart.charAt(i))) {
                                desStartNum += desStart.charAt(i);
                            }
                        }

                        if (partStartNum.equalsIgnoreCase(desStartNum) && !partStartLet.equalsIgnoreCase(desStartLet)) { //checks if the two inputs are on same row

                            /*Counts number of columns*/
                            for (int i = 0; i < 10 || i < rows; i++) {
                                row = sheet.getRow(i);
                                if (row != null) {
                                    tmp = sheet.getRow(i).getPhysicalNumberOfCells();
                                    if (tmp > cols) {
                                        cols = tmp;
                                    }
                                }
                            }

                            int findCount = 0; //used to count number of times a starting cell is found
                            for (int r = 0; r < rows && findCount < 2; r++) { //loops through all rows, stops when 2 starting cells are found
                                row = sheet.getRow(r); //gets row
                                if (row != null) { //makes sure row is not null to avoid null pointer exceptions
                                    for (int c = 0; c < cols; c++) { //goes through all columns
                                        cell = row.getCell(c); //gets cell from current row at current column index
                                        if (cell != null) { //makes sure cell is not null to avoid null pointer exceptions
                                            CellRangeAddress range = new CellRangeAddress(cell.getRowIndex(), cell.getRowIndex(), cell.getColumnIndex(), cell.getColumnIndex()); //gets the range address of current cell
                                            String rangeString = range.toString(); //converts address to string
                                            if (rangeString.indexOf(partStart) != -1) { //checks if the address is the same as the starting part number cell
                                                startingPartCell = cell; //sets data for starting part number cell
                                                findCount++; //increments the counter by one
                                            } else if (rangeString.indexOf(desStart) != -1) { //checks if the address is the same as the starting designator cell
                                                startDesCell = cell; //sets data for starting designator cell
                                                findCount++; //increments the counter by one
                                            }
                                        }
                                    }
                                }
                            }

                            if (startingPartCell != null && startDesCell != null) { //makes sure both starting cells exist
                                for (int r = startingPartCell.getRowIndex(); r < rows; r++) { //goes through sheet, starting from row that starting part number cell is on and ends on row user's input

                                    row = sheet.getRow(r); //gets row

                                    partCell = row.getCell(startingPartCell.getColumnIndex()); //gets first part number cell to be compared
                                    desCel = row.getCell(startDesCell.getColumnIndex()); //gets first designator cell to be compared


                                    partCell.setCellType(Cell.CELL_TYPE_STRING); //sets cell type
                                    String partCellCon = partCell.getStringCellValue(); //gets the string value of the cell
                                    desCel.setCellType(Cell.CELL_TYPE_STRING); //sets cell type
                                    String desCellCon = desCel.getStringCellValue(); //gets the string value of the cell

                                    String cvsSplit = ", "; //default characters used to split the cvs file's line into an array
                                    if (desCellCon.indexOf(", ") == -1) { //checks if the line contains the default splitter
                                        cvsSplit = ","; //if not found, make this the new splitter
                                    }

                                    String[] desLine = desCellCon.split(cvsSplit); //splits the line into an array
                                    LinkedList<String> designatorList = new LinkedList<String>(); //list of designators

                                    for (int i = 0; i < desLine.length; i++) { //goes through each element in array
                                        int lettersFound = 0; //counts number of letters found
                                        for (int l = 0; l < desLine[i].length(); l++) { //loops through each characer in the element
                                            if (Character.isLetter(desLine[i].charAt(l))) { //checks if the current character is a letter
                                                lettersFound++; //increments the letter count by 1
                                            }
                                        }
                                        if (lettersFound < 1) { //if an element is only a number, occurs when "C1,2"
                                            //This will get the last element, which will always be correct and get the letters from this and add it to the element with only numbers
                                            String fixedDes = ""; //initializes new designator
                                            boolean stay = true; //boolean to check if loop should keep running
                                            for (int l = 0; l < desLine[i - 1].length() && stay; l++) { //goes through each character in last element and while the stay boolean is true
                                                if (Character.isLetter(desLine[i - 1].charAt(l))) { //checks if the current character from the last element is a letter
                                                    fixedDes += desLine[i - 1].charAt(l); //if it is a letter, add it to new designator
                                                } else {
                                                    stay = false; //if its not a letter, stop running
                                                }
                                            }
                                            fixedDes += desLine[i]; //adds the numbers to the letters
                                            desLine[i] = fixedDes; //removes old element and adds fixed element
                                        }
                                        designatorList.add(desLine[i]); //adds fixed designator to designator list
                                    }

                                    for (int i = 0; i < designatorList.size(); i++) { //goes through all elements from designator list
                                        String letters = ""; //letters of designators
                                        String numbers = ""; //numbers of designator
                                        String finalNumbers = ""; //numbers used when there is a dash Ex "C1-8" or "C1-C8"
                                        String currentDesignator = designatorList.get(i); //designator to be checked
                                        if (currentDesignator.indexOf('-') != -1) { //if it has dashes
                                            int count = 0; //used to go through each character
                                            //Example: "C1-C9" First loop will get letters, Second loop will get numbers by looping until a dash is found, Third loop is used to increase the counter until a number is found
                                            while (Character.isLetter(currentDesignator.charAt(count))) { //checks if the current character is a letter, if not exits loop
                                                letters += currentDesignator.charAt(count); //adds to letters
                                                count++; //increments counter by 1
                                            }
                                            while (currentDesignator.charAt(count) != '-') { //loops until current character is a dash
                                                numbers += currentDesignator.charAt(count); //adds to the numbers
                                                count++; //increments counter by 1
                                            }
                                            while (!Character.isDigit(currentDesignator.charAt(count))) { //loops until a number is found
                                                count++; //increases count
                                            }
                                            finalNumbers = currentDesignator.substring(count); //gets final numbers by getting rest of string
                                            designatorList.set(i, currentDesignator.substring(0, currentDesignator.indexOf('-'))); //sets the current element. Changes "C1-C8" to just "C1"
                                            int firstNum = Integer.parseInt(numbers); //gets first number
                                            int secondNum = Integer.parseInt(finalNumbers); //gets second number, in examples case first number would be 1, second would be 8

                                            for (int k = 1; k <= secondNum - firstNum; k++) { //loops through the numbers from first number to second number. Skips first numebr because it was already set above
                                                i++; //increases counter so element can be added to specific position
                                                designatorList.add(i, letters + (firstNum + k)); //adds fixed designator to designator list at a specific spot by adding letters and adding this loop's counter to the starting number
                                            }
                                        }
                                    }

                                    String[] newArray = new String[designatorList.size()]; //makes an array the same size of the designator list
                                    for (int i = 0; i < designatorList.size(); i++) { //goes through each element of array
                                        newArray[i] = designatorList.get(i); //sets each element from list into array
                                    }

                                    desLine = newArray; //sets old array to this fixed array

                                    for (int a = 0; a < desLine.length; a++) { //loop through each element in array
                                        BufferedReader br = new BufferedReader(new FileReader(fileTwo)); //makes buffered reader to go through csv file

                                        String des = desLine[a]; //current designator to be used from designator array from Excel file

                                        timesFound = 0; //keeps track of times designator is found
                                        String line; //variable used to read line from CSV file
                                        String designator; //designator from CSV file

                                        while ((line = br.readLine()) != null) {


                                            designator = line.substring(0, line.indexOf(' '));
                                            /*partNumber = line.substring(line.indexOf(' ')); //full string with spaces
                                            String tempPart = "";

                                            tempPart = removeSpaces(partNumber);


                                            tempPart = tempPart.substring(tempPart.indexOf(' '));


                                            tempPart = removeSpaces(tempPart);


                                            partNumber = tempPart.substring(0, tempPart.indexOf(' '));*/


                                            if (des.equalsIgnoreCase(designator)) {
                                                updated.add(line);

                                                timesFound++;
                                            }
                                        }


                                        if (timesFound == 0) {

                                            output.println(des + " Not Found");
                                        } else if (timesFound > 1) {

                                            output.println(des + " More than one Found");
                                        } else if (timesFound == 1) {


                                            String lineRead = updated.get(updated.size() - 1); //gets last item

                                            if (lineRead.indexOf(des) == -1) {

                                                newList.add(lineRead);
                                            } else {

                                                int firstCharCount = 0;
                                                boolean addChar = true;
                                                int charPlace = 0;
                                                int lastCharPlace = 0;

                                                for (int l = 0; l < lineRead.length() && firstCharCount != 3; l++) {
                                                    if (lineRead.charAt(l) != ' ' && addChar) {
                                                        firstCharCount++;
                                                        addChar = false;
                                                    } else if (lineRead.charAt(l) == ' ') {
                                                        addChar = true;
                                                    }
                                                    if (firstCharCount == 2) {
                                                        charPlace = l + 1;
                                                    }

                                                }


                                                boolean addChara = false;
                                                String tempo = lineRead.substring(charPlace);
                                                boolean keepg = true;

                                                for (int l = 0; l < tempo.length() && keepg; l++) {
                                                    if (tempo.charAt(l) != ' ' && addChara) {
                                                        lastCharPlace = l;
                                                        keepg = false;
                                                    } else if (tempo.charAt(l) == ' ') {
                                                        addChara = true;
                                                    }
                                                }


                                                String partSpace = partCellCon;


                                                for (int le = 0; le < 40 && partSpace.length() < 30; le++) {
                                                    partSpace += ' ';
                                                }


                                                newList.add(lineRead.substring(0, charPlace) + partSpace + lineRead.substring(lastCharPlace + charPlace));


                                            }


                                        }
                                    }


                                }
                            } else {
                                JOptionPane.showMessageDialog(null, "Cell Entered is Invald");
                            }
                        } else {
                            JOptionPane.showMessageDialog(null, "Cells are not in same row");
                        }

                        try {


                            File newFile = new File(file.getParentFile().getAbsolutePath() + File.separator + "Updated.csv");
                            FileOutputStream fo = new FileOutputStream(newFile);


                            PrintWriter pw = new PrintWriter(fo);

                            for (int p = 0; p < newList.size(); p++) {
                                pw.println(newList.get(p));
                            }

                            pw.close();
                            output.close();
                            Runtime rt = Runtime.getRuntime();
                            String txtPath = fileTwo.getParentFile().getAbsolutePath() + File.separator + "Differences.txt";
                            Process p = rt.exec("notepad " + txtPath);

                            Runtime rtTwo = Runtime.getRuntime();
                            String txtPathTwo = file.getParentFile().getAbsolutePath() + File.separator + "Updated.csv";
                            Process pTwo = rtTwo.exec("notepad " + txtPathTwo);


                        } catch (Exception excep) {
                            excep.printStackTrace();
                        }
                    } catch (Exception exe) {
                        exe.printStackTrace();
                        JOptionPane.showMessageDialog(null, "Error");
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Wrong File Format");
                }
            } else {
                JOptionPane.showMessageDialog(null, "Invalid File");
            }

        }
    }

    public static String removeSpaces(String designator) {
        //get rid of spaces
        String tempDes = "";
        boolean stay = true;
        boolean keep = true;
        for (int i = 0; i < designator.length(); i++) {
            if (designator.charAt(i) != ' ' || keep == false) {
                tempDes += designator.charAt(i);
                stay = false;
            } else {
                if (stay == false) {
                    keep = false;
                }
            }

        }
        return tempDes;
    }


    public class FileOneButListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            FileChooserGUI fcg;
            if (directory != null) {
                fcg = new FileChooserGUI(directory, "excel");
            } else {
                fcg = new FileChooserGUI("excel");
            }
            fileOneString = fcg.getPath();
            fileOneField.setText(fileOneString);
            directory = fcg.getDir();
        }
    }

    public class FileTwoButListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            FileChooserGUI fcg;

            if (directory != null) {
                fcg = new FileChooserGUI(directory, "csv");
            } else {
                fcg = new FileChooserGUI("csv");
            }
            fileTwoString = fcg.getPath();
            fileTwoField.setText(fileTwoString);
            directory = fcg.getDir();
        }
    }

    public class HelpButtonListener implements ActionListener{

        /**
         * Invoked when an action occurs.
         *
         * @param e
         */
        @Override
        public void actionPerformed(ActionEvent e) {
            final String newline = "\n";
            area = new JTextArea(20, 20);
            area.setEditable(false);
            area.append("AMSOutput.jar" + newline);
            area.append("------------------------------------------------------------------------------------------------------------------------------" + newline);
            area.append("Gets part number and designator from Excel file and looks for the same designator in the CSV file." + newline);
            area.append("If it is found, the program will change the CSV file's part number and replace it with the part number found in the Excel file."+newline);
            area.append("When it is done, the updated file will be saved in the same directory as the 2 files entered and will automatically open." + newline);
            panel.add(area);
            panel.revalidate();
            panel.repaint();
            helpButton.setEnabled(false);
        }
    }
}
