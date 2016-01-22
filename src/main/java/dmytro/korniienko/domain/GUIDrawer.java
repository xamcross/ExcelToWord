package dmytro.korniienko.domain;

import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;

public class GUIDrawer extends JFrame {

	private final static String GUI_SCREEN_TITLE_TEXT = "ExcelToWord 9000";
	private final static String GUI_LABEL_EXCEL_FILE_URI_TEXT = "Please, choose .xlsx Excel file to get data from:";
	private final static String GUI_LABEL_WORD_FILE_URI_TEXT = "Please, choose .docx Word file to copy data to:";
	private final static String GUI_BUTTON_BROWSE_EXCEL_FILE_TEXT = "Find Excel file...";
	private final static String GUI_BUTTON_BROWSE_WORD_FILE_TEXT = "Find Word file...";
	private final static String GUI_BUTTON_LAUNCH = "Launch";

	private final static String FILTER_EXCEL_LABEL_TEXT = ".xlsx";
	private final static String FILTER_WORD_LABEL_TEXT = ".docx";
	private final static String FILTER_EXCEL_EXTENSION = "xlsx";
	private final static String FILTER_WORD_EXTENSION = "docx";

	private final static int GUI_SCREEN_WIDTH_SMALL = 640;
	private final static int GUI_SCREEN_HEIGHT_SMALL = 480;
	private final static int GUI_SCREEN_WIDTH_MEDIUM = 800;
	private final static int GUI_SCREEN_HEIGHT_MEDIUM = 600;
	private final static int GUI_SCREEN_WIDTH_LARGE = 1024;
	private final static int GUI_SCREEN_HEIGHT_LARGE = 768;

	private final static int GUI_TEXT_FIELD_EXCEL_FILE_WIDTH = 50;
	private final static int GUI_TEXT_FIELD_WORD_FILE_WIDTH = 50;

	private static JLabel labelExcelFileUri;
	private static JLabel labelWordFileUri;
	private static JButton btnBrowseExcelFile;
	private static JButton btnBrowseWordFile;
	private static JButton btnLaunch;
	private static JTextField tfExcelFileUri;
	private static JTextField tfWordFileUri;
	private static JFileChooser fcBrowseFile;

	private static GUIDrawer instance;

	private String excelUri;

	private String wordUri;

	public GUIDrawer() {
		instance = this;
		setLayout(new FlowLayout());

		initializeInputControls();

		addInputControls();

		assignActionListeners();

		drawBasicWindow();
	}

	private void assignActionListeners() {

		btnBrowseExcelFile.addActionListener(new OnClickBrowseListener(tfExcelFileUri));

		btnBrowseWordFile.addActionListener(new OnClickBrowseListener(tfWordFileUri));

		btnLaunch.addActionListener(new OnClickLaunchListener());
	}

	private void addInputControls() {
		this.add(labelExcelFileUri);
		this.add(tfExcelFileUri);
		this.add(btnBrowseExcelFile);
		this.add(labelWordFileUri);
		this.add(tfWordFileUri);
		this.add(btnBrowseWordFile);
		this.add(btnLaunch);

	}

	private void drawBasicWindow() {
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		this.setSize(GUI_SCREEN_WIDTH_MEDIUM, GUI_SCREEN_HEIGHT_MEDIUM);
		this.setVisible(true);
		this.setTitle(GUI_SCREEN_TITLE_TEXT);

	}

	private void initializeInputControls() {
		labelExcelFileUri = new JLabel(GUI_LABEL_EXCEL_FILE_URI_TEXT);
		labelWordFileUri = new JLabel(GUI_LABEL_WORD_FILE_URI_TEXT);

		btnBrowseExcelFile = new JButton(GUI_BUTTON_BROWSE_EXCEL_FILE_TEXT);
		btnBrowseWordFile = new JButton(GUI_BUTTON_BROWSE_WORD_FILE_TEXT);
		btnLaunch = new JButton(GUI_BUTTON_LAUNCH);

		tfExcelFileUri = new JTextField(GUI_TEXT_FIELD_EXCEL_FILE_WIDTH);
		tfWordFileUri = new JTextField(GUI_TEXT_FIELD_WORD_FILE_WIDTH);

		fcBrowseFile = new JFileChooser();

	}

	class OnClickBrowseListener implements ActionListener {

		private JTextField textField;
		private FileFilter filter;

		public OnClickBrowseListener(JTextField textField) {
			this.textField = textField;
		}

		@Override
		public void actionPerformed(ActionEvent e) {
			JButton clickedButton = (JButton) e.getSource();
			if (clickedButton.getText().equals(GUI_BUTTON_BROWSE_EXCEL_FILE_TEXT)) {
				filter = new FileNameExtensionFilter(FILTER_EXCEL_LABEL_TEXT, FILTER_EXCEL_EXTENSION);
			} else if (clickedButton.getText().equals(GUI_BUTTON_BROWSE_WORD_FILE_TEXT)) {
				filter = new FileNameExtensionFilter(FILTER_WORD_LABEL_TEXT, FILTER_WORD_EXTENSION);
			}
			fcBrowseFile.resetChoosableFileFilters();
			if (filter != null) {
				fcBrowseFile.setAcceptAllFileFilterUsed(false);
				fcBrowseFile.setFileFilter(filter);
			}
			fcBrowseFile.showOpenDialog(instance);
			File file = fcBrowseFile.getSelectedFile();

			if (file != null) {
				textField.setText(file.getName());
				fcBrowseFile.setSelectedFile(null);
			}
		}

	}

	class OnClickLaunchListener implements ActionListener {

		@Override
		public void actionPerformed(ActionEvent e) {
			excelUri = tfExcelFileUri.getText();
			wordUri = tfWordFileUri.getText();
			ExcelParser parser = new ExcelParser(excelUri, wordUri);
			parser.parseExcelProjectNames();
		}

	}

}
