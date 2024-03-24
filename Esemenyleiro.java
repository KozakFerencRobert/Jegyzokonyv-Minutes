package Jegyzokonyv;

//import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.io.File;
//import java.io.FileInputStream;
import java.io.FileOutputStream;
//import java.io.IOException;
//import java.io.InputStream;
import java.math.BigInteger;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class Esemenyleiro extends JFrame {
	private static final long serialVersionUID = 1L;
	private JLabel locationLabel, timeLabel, descriptionLabel, actionLabel, attachmentLabel;
	private JTextField locationField, timeField, actionField, attachmentField;

	JTextArea descriptionField;
	private JButton createButton, exitButton;

	// Az előre elkészített táblázat fejléc adatai
	private static final String[][] HEADER_DATA = { { "", "", "" }, { "", "", "" }, { "", "", "" } // {"","céginfó",""}
	};

	// A táblázatba beillesztendő képek elérési útvonalai
	/*
	 * private static final String[] IMAGES_PATH = {
	 * //"D:/Eclipse_workspace/Jegyzokonyv_1/src/Jegyzokonyv/companylogo.jpg",
	 * "D:/Eclipse_workspace/Jegyzokonyv_1/src/Jegyzokonyv/golden.jpg" //cég logó
	 * elérési útvonal };
	 */
	private void createHeaderTable(XWPFDocument document) {
		XWPFTable table = document.createTable(1, 3);

		// Képek beillesztése a táblázatba (bal oldal és jobb oldal)
		// insertImage(table, 0, IMAGES_PATH[0], 140, 100,ParagraphAlignment.LEFT);
		// insertImage(table, 2, IMAGES_PATH[1], 140, 100, ParagraphAlignment.RIGHT);

		// Középső cella a céginformációhoz
		XWPFTableCell centerCell = table.getRow(0).getCell(1);
		centerCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
		XWPFParagraph paragraph = centerCell.getParagraphs().get(0);
		XWPFRun run = paragraph.createRun();
		run.setText(HEADER_DATA[2][1]);
		run.setFontSize(12);
		paragraph.setAlignment(ParagraphAlignment.CENTER);

		// Táblázat formázása
		CTTbl ctTbl = table.getCTTbl();
		CTTblPr tblPr = ctTbl.getTblPr();
		if (tblPr == null)
			tblPr = ctTbl.addNewTblPr();
		tblPr.addNewTblW().setW(BigInteger.valueOf(8000)); // Táblázat szélessége (8000 = 100%)

		for (int i = 0; i < 3; i++) {
			table.getRow(0).getCell(i).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(2666)); // Oszlop
																											// szélessége
																											// (2666 =
																											// 33.33%)
		}
	}

	/*
	 * private void insertImage(XWPFTable table, int cellIndex, String imagePath,
	 * int width, int height, ParagraphAlignment alignment) { XWPFTableCell cell =
	 * table.getRow(0).getCell(cellIndex); try { File imageFile = new
	 * File(imagePath); InputStream fis = new FileInputStream(imageFile); int
	 * imageType = XWPFDocument.PICTURE_TYPE_JPEG; String imageName =
	 * imageFile.getName();
	 * 
	 * XWPFParagraph paragraph = cell.getParagraphs().get(0);
	 * paragraph.setAlignment(alignment); XWPFRun run = paragraph.createRun();
	 * run.addBreak(); run.addPicture(fis, imageType, imageName, Units.toEMU(width),
	 * Units.toEMU(height)); } catch (Exception e) { e.printStackTrace(); } }
	 */
	public Esemenyleiro() {
		// Ablak címe
		setTitle("Eseményleíró Jegyzőkönyv");
		// Ablak mérete
		setSize(1500, 750);
		// Az ablak bezárásakor álljon le az alkalmazás
		setDefaultCloseOperation(EXIT_ON_CLOSE);

		Font AreaFont = new Font("Arial", Font.BOLD, 12);

		// Felhasználói felület létrehozása
		JPanel panel = new JPanel();
		panel.setLayout(new GridLayout(6, 2, 10, 10));
		panel.setBackground(new Color(217,217,214));

		locationLabel = new JLabel("	Esemény helye:");
		locationField = new JTextField(20);

		timeLabel = new JLabel("	Esemény ideje (év/hónap/nap-óra:perc): ");
		timeField = new JTextField(20);

		descriptionLabel = new JLabel("		Eseményleírás:");
		// descriptionField = new JTextArea(4,5);

		descriptionField = new JTextArea();
		descriptionField.setFont(AreaFont);
		descriptionField.setLineWrap(true); // Engedélyezi az automatikus sorváltást
		descriptionField.setWrapStyleWord(true); // Csak szóköznél vágja el a sort

		// Korlátozza a sorok számát
		descriptionField.addKeyListener(new KeyListener() {
			@Override
			public void keyTyped(KeyEvent e) {
				if (descriptionField.getLineCount() >= 5) { // Maximum 5 sor
					e.consume(); // Elnyomja a billentyű leütést
				}
			}

			@Override
			public void keyPressed(KeyEvent e) {
			}

			@Override
			public void keyReleased(KeyEvent e) {
			}
		});

		actionLabel = new JLabel("		Intézkedés:");
		actionField = new JTextField(20);

		attachmentLabel = new JLabel("		A jegyzőkönyvhöz mellékelve:");
		attachmentField = new JTextField(20);

		createButton = new JButton("Létrehozás");
		createButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				createWordFile();
			}
		});

		exitButton = new JButton("Kilépés");
		exitButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				dispose(); // Az ablak bezárása
			}
		});
		panel.add(locationLabel);
		panel.add(locationField);
		panel.add(timeLabel);
		panel.add(timeField);
		panel.add(descriptionLabel);
		JScrollPane descriptionScrollPane = new JScrollPane(descriptionField);
		descriptionScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
		panel.add(descriptionScrollPane);
		// panel.add(descriptionField);
		panel.add(actionLabel);
		panel.add(actionField);
		panel.add(attachmentLabel);
		panel.add(attachmentField);
		panel.add(createButton);
		panel.add(exitButton);

		createButton.setForeground(Color.BLUE);
		exitButton.setForeground(Color.RED);

		Font FontSize = new Font("Arial", Font.BOLD, 18);

		createButton.setFont(FontSize);
		exitButton.setFont(FontSize);

		for (Component component : panel.getComponents()) {
			if (component instanceof JTextField) {
				JTextField textField = (JTextField) component;
				textField.setFont(FontSize);
			}
		}
		for (Component component : panel.getComponents()) {
			if (component instanceof JLabel) {
				JLabel label = (JLabel) component;
				label.setFont(FontSize);
			}
		}

		add(panel);
	}

	public static boolean isValidDateFormat(String inputDate, String format) {
		SimpleDateFormat sdf = new SimpleDateFormat(format);
		sdf.setLenient(false);

		try {
			sdf.parse(inputDate);
			return true;
		} catch (ParseException e) {
			return false;
		}
	}

	private void createWordFile() {

		// Dátum és idő felolvasása a mezőkből
		String location = locationField.getText();
		String time = timeField.getText();
		String description = descriptionField.getText();
		String action = actionField.getText();
		String attachment = attachmentField.getText();
		String format = "yyyy/MM/dd-HH:mm";

		if (isValidDateFormat(time, format) && !location.isEmpty() && !description.isEmpty() && !action.isEmpty()) {
			try (// Új Word dokumentum létrehozása

					XWPFDocument document = new XWPFDocument()) {

				// Dátum és idő formázása a fájlnévhez
				LocalDateTime currentDateTime = LocalDateTime.now();
				DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");
				String formattedDateTime = currentDateTime.format(formatter);

				String folderName = "Jegyzokonyv";
				File folder = new File(folderName);
				if (!folder.exists()) {
					folder.mkdir();
				}
				// Címsor hozzáadása
				createHeaderTable(document);

				XWPFParagraph titleParagraph = document.createParagraph();
				titleParagraph.setAlignment(ParagraphAlignment.CENTER);
				XWPFRun titleRun = titleParagraph.createRun();
				titleRun.setText("Eseményleíró Jegyzőkönyv");
				titleRun.setBold(true);
				titleRun.setFontSize(20);

				// Esemény rész hozzáadása

				XWPFParagraph eventParagraph = document.createParagraph();
				eventParagraph.setAlignment(ParagraphAlignment.LEFT);
				XWPFRun eventRun = eventParagraph.createRun();
				eventRun.setText("Esemény helye:  " + location);
				eventRun.addBreak();
				eventRun.setText("Esemény ideje (év/hónap/nap-óra:perc):  " + time);
				eventRun.addBreak();
				eventRun.setText("Eseményleírás:  " + description);
				eventRun.addBreak();
				eventRun.setText("Intézkedés:  " + action);

				// Jegyzőkönyvhöz mellékelt dokumentum hozzáadása
				XWPFParagraph attachmentParagraph = document.createParagraph();
				attachmentParagraph.setAlignment(ParagraphAlignment.LEFT);
				XWPFRun attachmentRun = attachmentParagraph.createRun();
				attachmentRun.setText("A jegyzőkönyvhöz mellékelve:  " + attachment);

				XWPFParagraph kmfParagraph = document.createParagraph();
				kmfParagraph.setAlignment(ParagraphAlignment.LEFT);
				XWPFRun kmfRun = kmfParagraph.createRun();
				kmfRun.setText("                                                          K.m.f.");

				// Tanúk rész hozzáadása
				XWPFParagraph witnessesParagraph = document.createParagraph();
				witnessesParagraph.setAlignment(ParagraphAlignment.LEFT);
				XWPFRun witnessesRun = witnessesParagraph.createRun();

				witnessesRun.setText("Tanúk:");
				witnessesRun.addBreak();
				for (int i = 0; i < 5; i++) {
					witnessesRun.setText("              ___________________________");
					witnessesRun.addBreak();
				}

				// Eseményt leíró aláírás hozzáadása
				XWPFParagraph signatureParagraph = document.createParagraph();
				signatureParagraph.setAlignment(ParagraphAlignment.LEFT);
				XWPFRun signatureRun = signatureParagraph.createRun();
				signatureRun.setText("Eseményt leíró aláírása: __________________________");

				// Word fájl mentése a mappába
				String fileName = folderName + "/" + "Jegyzokonyv_" + formattedDateTime + ".docx";
				FileOutputStream out = new FileOutputStream(new File(fileName));
				// Clear text from the input fields after processing and creating the Word file
				locationField.setText("");
				timeField.setText("");
				descriptionField.setText("");
				actionField.setText("");
				attachmentField.setText("");

				document.write(out);
				out.close();

				JOptionPane.showMessageDialog(this, "A jegyzőkönyv sikeresen létrehozva:\n" + fileName);
			} catch (Exception e1) {
				e1.printStackTrace();
			}
		} else {
			JOptionPane.showMessageDialog(Esemenyleiro.this,
					"Hiba!!!!\nNincs megadva az esemény helye vagy ideje vagy hibás dátum formátum\nHelyes formátum:\n(év/hónap/nap-óra:perc");
		} /*
			 * catch (HeadlessException e) { e.printStackTrace(); }
			 */
	}

	public static void main(String[] args) {
		SwingUtilities.invokeLater(() -> new Esemenyleiro().setVisible(true));
	}
}