package AMSOutput;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.File;

/**
 * Created by rahul on 2017-06-04.
 */
public class FileChooserGUI extends JPanel {
    static private final String newline = "\n";
    JTextArea log;
    JFileChooser fc;
    private String path = "";
    private File dir;
    private String typeFile;

    public String getPath(){
        return path;
    }

    public File getDir(){
        return dir;
    }


    public FileChooserGUI(String type) {
        super(new BorderLayout());
        this.typeFile = type;

        //Create the log first, because the action listeners
        //need to refer to it.
        log = new JTextArea(5,20);
        log.setMargin(new Insets(5,5,5,5));
        log.setEditable(false);
        JScrollPane logScrollPane = new JScrollPane(log);

        //Create a file chooser
        fc = new JFileChooser();
        if (type.equalsIgnoreCase("csv")){
            FileNameExtensionFilter filter = new FileNameExtensionFilter("CSV Files", "csv");
            fc.setFileFilter(filter);
        }else{
            FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel File", "xls", "xlsx");
            fc.setFileFilter(filter);
        }

        int returnVal = fc.showOpenDialog(FileChooserGUI.this);

        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File file = fc.getSelectedFile();
            dir = fc.getCurrentDirectory();
            path = file.getAbsolutePath();
            //This is where a real application would open the file.
            log.append("Opening: " + file.getName() + "." + newline);
        } else {
            log.append("Open command cancelled by user." + newline);
        }
        log.setCaretPosition(log.getDocument().getLength());




    }

    public FileChooserGUI(File directory, String type) {
        super(new BorderLayout());
        this.typeFile = type;

        //Create the log first, because the action listeners
        //need to refer to it.
        log = new JTextArea(5,20);
        log.setMargin(new Insets(5,5,5,5));
        log.setEditable(false);
        JScrollPane logScrollPane = new JScrollPane(log);

        //Create a file chooser
        fc = new JFileChooser();
        fc.setCurrentDirectory(directory);
        if (type.equalsIgnoreCase("csv")){
            FileNameExtensionFilter filter = new FileNameExtensionFilter("CSV Files", "csv");
            fc.setFileFilter(filter);
        }else{
            FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel File", "xls", "xlsx");
            fc.setFileFilter(filter);
        }

        int returnVal = fc.showOpenDialog(FileChooserGUI.this);

        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File file = fc.getSelectedFile();
            dir = fc.getCurrentDirectory();
            path = file.getAbsolutePath();
            //This is where a real application would open the file.
            log.append("Opening: " + file.getName() + "." + newline);
        } else {
            log.append("Open command cancelled by user." + newline);
        }
        log.setCaretPosition(log.getDocument().getLength());




    }




    /**
     * Create the GUI and show it.  For thread safety,
     * this method should be invoked from the
     * event dispatch thread.

    private static void createAndShowGUI() {
        //Create and set up the window.
        JFrame frame = new JFrame("FileChooserDemo");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        //Add content to the window.
        frame.add(new FileChooserGUI(this.typeFile));

        //Display the window.
        frame.pack();
        frame.setVisible(true);
    }*/

}

