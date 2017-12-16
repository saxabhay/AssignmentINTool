package AssignmentIN;



import java.awt.Color;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.border.EmptyBorder;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class Assignment extends JFrame implements ActionListener {
	private static final long serialVersionUID = 142149479884157820L;
	private JPanel contentPanel;
	private JTextField emailIDTextField;
	JLabel emailLabel;
	JButton createButton;
	JButton closeButton;
	protected static String username;
	protected static String password;
	private JLabel nameLabel;
	private JTextField nameTextField;

	public static void main(String[] args)
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {

					Assignment account = new Assignment();
					account.setVisible(true);

				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public Assignment() {

		setTitle("IN Assignment Made By ABHAY SAXENA");
		setResizable(false);
		setDefaultCloseOperation(3);
		setBounds(100, 100, 450, 450);

		contentPanel = new JPanel();
		contentPanel.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPanel);
		contentPanel.setLayout(null);

		emailLabel = new JLabel("Crowd Password");
		emailLabel.setFont(new Font("Georgia", Font.BOLD | Font.ITALIC, 14));
		emailLabel.setBackground(new Color(204, 204, 204));
		emailLabel.setBounds(71, 133, 150, 17);
		contentPanel.add(emailLabel);

		emailIDTextField = new JTextField();
		emailIDTextField.setFont(new Font("Sylfaen", Font.ITALIC, 16));
		emailIDTextField.setBounds(200, 125, 200, 28);
		contentPanel.add(emailIDTextField);
		emailIDTextField.setColumns(10);

		createButton = new JButton("Submit");
		createButton.setBounds(71, 300, 153, 44);
		contentPanel.add(createButton);
		createButton.addActionListener(this);

		closeButton = new JButton("Cancel");
		closeButton.setBounds(280, 300, 136, 44);
		contentPanel.add(closeButton);

		nameLabel = new JLabel("Username");
		nameLabel.setFont(new Font("Georgia", Font.BOLD | Font.ITALIC, 15));
		nameLabel.setBackground(new Color(204, 204, 204));
		nameLabel.setBounds(71, 42, 109, 36);
		contentPanel.add(nameLabel);

		nameTextField = new JTextField();
		nameTextField.setFont(new Font("Sylfaen", Font.ITALIC, 16));
		nameTextField.setColumns(10);
		nameTextField.setBounds(200, 42, 200, 28);
		contentPanel.add(nameTextField);

		closeButton.addActionListener(this);

	}

	public void actionPerformed(ActionEvent event) {

		if (event.getSource() == closeButton) {
			super.dispose();
		}

		if (event.getSource() == createButton) {

			username = nameTextField.getText();
			password = emailIDTextField.getText();

			super.dispose();
		}

		System.setProperty("webdriver.chrome.driver",
				"/Users/saxabhay/eclipse-workspace/AssignmentTool/Assignment/driver/chromedriver");
		ChromeOptions options = new ChromeOptions();
		options.addArguments("window-size=1400,800");
		WebDriver driver = new ChromeDriver(options);

		driver.get(
				"https://wiki.labcollab.net/confluence/pages/viewpage.action?spaceKey=COPS&title=IN+Language+Skills+in+JIRA");
		driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
		driver.findElement(By.id("os_username")).sendKeys(username);
		driver.findElement(By.id("os_password")).sendKeys(password);
		driver.findElement(By.id("loginButton")).click();

		try {
			Thread.sleep(9000);
		} catch (InterruptedException e) {

			e.printStackTrace();
		}
		driver.findElement(By.xpath("(//a[text()='Refresh'])[1]")).click();
		driver.findElement(By.xpath("(//a[text()='Refresh'])[2]")).click();
		driver.findElement(By.xpath("(//a[text()='Refresh'])[3]")).click();

		try {
			Thread.sleep(4000);
		} catch (InterruptedException e1) {
			e1.printStackTrace();
		}
		FileInputStream f;
		Workbook wb = null;
		FileOutputStream fos1 = null;
		FileOutputStream fos2 = null;
		FileOutputStream fos3 = null;
		try {
			f = new FileInputStream(new File("/Users/saxabhay/Desktop/AssignmentIN.xlsx"));
			wb = WorkbookFactory.create(f);
			fos1 = new FileOutputStream(new File("/Users/saxabhay/Desktop/AssignmentIN.xlsx"));
			fos2 = new FileOutputStream("/Users/saxabhay/Desktop/AssignmentIN.xlsx");
			fos3 = new FileOutputStream("/Users/saxabhay/Desktop/AssignmentIN.xlsx");
		} catch (FileNotFoundException e) {

			e.printStackTrace();
		} catch (EncryptedDocumentException e) {

			e.printStackTrace();
		} catch (InvalidFormatException e) {

			e.printStackTrace();
		} catch (IOException e) {

			e.printStackTrace();
		}
		Sheet sh = wb.getSheet("FlashBriefing");
		List<WebElement> alltr = driver.findElements(By.xpath("(//tbody)[1]/tr"));// Content/Flash Briefing IN-English
																					// Skills
		System.out.println("Content Flash Briefing IN-English Skills");
		for (int i = 0; i < alltr.size(); i++) {
			Row row = sh.createRow(i);
			List<WebElement> alltd = alltr.get(i).findElements(By.xpath(".//td|.//th"));
			for (int j = 0; j < alltd.size(); j++) {
				String text = alltd.get(j).getText();
				System.out.print(text + " ");
				Cell cell = row.createCell(j);
				cell.setCellType(CellType.STRING);
				cell.setCellValue(text);

			}
			System.out.println();
		}
		try {
			wb.write(fos1);
			fos1.close();
		} catch (IOException e) {

			e.printStackTrace();
		}

		Sheet sh1 = wb.getSheet("IN-English");
		List<WebElement> alltr1 = driver.findElements(By.xpath("(//tbody)[2]/tr"));// Custom IN-English Skills
		System.out.println("Custom IN-English Skills");
		for (int i = 0; i < alltr1.size(); i++) {
			Row row = sh1.createRow(i);
			List<WebElement> alltd1 = alltr1.get(i).findElements(By.xpath(".//td|.//th"));
			for (int j = 0; j < alltd1.size(); j++) {
				String text = alltd1.get(j).getText();
				System.out.print(text + " ");
				Cell cell = row.createCell(j);
				cell.setCellType(CellType.STRING);
				cell.setCellValue(text);

			}
			System.out.println();
		}
		try {
			wb.write(fos2);
			fos2.close();
		} catch (IOException e) {

			e.printStackTrace();
		}

		Sheet sh2 = wb.getSheet("Smart Home");
		List<WebElement> alltr2 = driver.findElements(By.xpath("(//tbody)[3]/tr"));// Smart Home IN-English Skills
		System.out.println("Smart Home IN-English Skills");
		for (int i = 0; i < alltr2.size(); i++) {
			Row row = sh2.createRow(i);
			List<WebElement> alltd2 = alltr2.get(i).findElements(By.xpath(".//td|.//th"));
			for (int j = 0; j < alltd2.size(); j++) {
				String text = alltd2.get(j).getText();
				System.out.print(text + " ");
				Cell cell = row.createCell(j);
				cell.setCellType(CellType.STRING);
				cell.setCellValue(text);
			}
			System.out.println();
		}
		try {
			wb.write(fos3);
			fos3.close();
		} catch (IOException e) {

			e.printStackTrace();
		}
        System.out.println();
		driver.quit();
	}
}

