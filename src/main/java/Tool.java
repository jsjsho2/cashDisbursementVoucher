import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.*;

import javax.swing.*;
import javax.swing.border.CompoundBorder;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Map;

public class Tool extends JFrame {

    int frameWidth = 500;

    GridBagConstraints gbc = new GridBagConstraints();

    JFrame jFrame = new JFrame();
    JPanel mainPanel = new JPanel(new GridBagLayout());

    JTabbedPane writeCashDisbursementVoucherTap = new JTabbedPane();
    JPanel writeCashDisbursementVoucherfunctionBtn = new JPanel(new GridLayout(0, 3, 10, 0));

    Tool() {
        jFrame.setTitle("귀찮아서 만든 지출결의서 툴");
        System.out.println("툴 실행...");

        makeWriteCashDisbursementVoucherPanel();

        jFrame.setContentPane(mainPanel);
        jFrame.setPreferredSize(new Dimension(frameWidth, 380));
        jFrame.setMinimumSize(new Dimension(frameWidth, 300));
        jFrame.pack();
        jFrame.setVisible(true);
    }

    public void makeWriteCashDisbursementVoucherPanel() {
        writeCashDisbursementVoucherfunctionBtn.setVisible(true);

        JButton runBtn = new JButton("실행");

        writeCashDisbursementVoucherfunctionBtn.add(runBtn);
//        writeCashDisbursementVoucherfunctionBtn.setBorder(new EmptyBorder(10, 0, 10, 0));

        mainPanel.add(writeCashDisbursementVoucherfunctionBtn, setGbc(1, 1, 0, 0));
        mainPanel.add(writeCashDisbursementVoucherTap, setGbc(1, 10, 0, 1));

        JPanel selectExcelPanel = new JPanel(new GridBagLayout());

        JPanel selectExcelTypePanel = new JPanel();
        selectExcelTypePanel.setBorder(new CompoundBorder(new EmptyBorder(10, 10, 10, 10), BorderFactory.createTitledBorder("다운받은 EXCEL 내역 타입")));

        ButtonGroup refTypeGroup = new ButtonGroup();
        JRadioButton refTypeRadio[] = new JRadioButton[2];
        refTypeRadio[0] = new JRadioButton("교통카드이용내역");
        refTypeRadio[1] = new JRadioButton("결재예정금액");
        refTypeRadio[0].setSelected(true);

        refTypeGroup.add(refTypeRadio[0]);
        refTypeGroup.add(refTypeRadio[1]);
        selectExcelTypePanel.add(refTypeRadio[0]);
        selectExcelTypePanel.add(refTypeRadio[1]);

        selectExcelPanel.add(selectExcelTypePanel, setGbc(1, 1, 0, 0));

        JPanel refExcelPanel = new JPanel();
        refExcelPanel.setBorder(new CompoundBorder(new EmptyBorder(10, 10, 10, 10), BorderFactory.createTitledBorder("다운받은 EXCEL 내역")));

        JTextField refFilePath = new JTextField("", 22);
        JButton refSeachFilePath = new JButton("파일선택");

        refExcelPanel.add(refFilePath);
        refExcelPanel.add(refSeachFilePath);

        selectExcelPanel.add(refExcelPanel, setGbc(1, 1, 0, 1));

        JPanel companyExcelPanel = new JPanel();
        companyExcelPanel.setBorder(new CompoundBorder(new EmptyBorder(10, 10, 10, 10), BorderFactory.createTitledBorder("지출결의서")));

        JTextField companyFilePath = new JTextField("", 22);
        JButton companySeachFilePath = new JButton("파일선택");

        companyExcelPanel.add(companyFilePath);
        companyExcelPanel.add(companySeachFilePath);

        selectExcelPanel.add(companyExcelPanel, setGbc(1, 1, 0, 2));

        JScrollPane dbInfoScroll = new JScrollPane(selectExcelPanel);
        dbInfoScroll.setBorder(BorderFactory.createLineBorder(Color.LIGHT_GRAY));

        writeCashDisbursementVoucherTap.add("엑셀선택", dbInfoScroll);

        refSeachFilePath.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                File f = new File(".");
                String rootPath = String.valueOf(f.getAbsoluteFile());
                rootPath = rootPath.substring(0, rootPath.length() - 1);
                final JFileChooser fileDialog = new JFileChooser(rootPath);

                int returnVal = fileDialog.showOpenDialog(mainPanel);

                if (returnVal == JFileChooser.APPROVE_OPTION) {
                    File file = fileDialog.getSelectedFile();
                    refFilePath.setText(file.toString());
                }
            }
        });

        companySeachFilePath.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                File f = new File(".");
                String rootPath = String.valueOf(f.getAbsoluteFile());
                rootPath = rootPath.substring(0, rootPath.length() - 1);
                final JFileChooser fileDialog = new JFileChooser(rootPath);

                int returnVal = fileDialog.showOpenDialog(mainPanel);

                if (returnVal == JFileChooser.APPROVE_OPTION) {
                    File file = fileDialog.getSelectedFile();
                    companyFilePath.setText(file.toString());
                }
            }
        });

        runBtn.addActionListener(e -> {
            int result = JOptionPane.showConfirmDialog(null, "실행하시겠습니까?");

            if(result == 0){
                ArrayList rowValues = new ArrayList();
                String importFilePath = "C:\\Users\\JSHPC\\Desktop\\교통카드이용내역.xls";//refFilePath.getText();
                HSSFWorkbook targetWorkbook = null;

                try {
                    FileInputStream fis = new FileInputStream(importFilePath);
                    targetWorkbook = new HSSFWorkbook(fis);
                } catch (Exception ex) {
                    ex.printStackTrace();
                }

                HSSFSheet targetSheet = targetWorkbook.getSheetAt(0);
                int rows = targetSheet.getPhysicalNumberOfRows();

                for (int rowindex = 1; rowindex <= rows; rowindex++) {
                    HSSFRow row = targetSheet.getRow(rowindex);
                    ArrayList rowValue = new ArrayList();

                    if (row != null) {
                        HSSFCell cell = row.getCell(0);
                        rowValue.add(cell.getStringCellValue());

                        cell = row.getCell(2);
                        rowValue.add(cell.getStringCellValue());

                        cell = row.getCell(3);
                        rowValue.add(cell.getStringCellValue());

                        cell = row.getCell(4);
                        rowValue.add(cell.getStringCellValue());

                        cell = row.getCell(5);
                        rowValue.add(cell.getStringCellValue());

                        cell = row.getCell(6);
                        rowValue.add(cell.getStringCellValue());

                        cell = row.getCell(7);
                        rowValue.add(cell.getStringCellValue());

                        cell = row.getCell(9);
                        rowValue.add(cell.getNumericCellValue());

                        cell = row.getCell(10);
                        rowValue.add(cell.getStringCellValue());

                        rowValues.add(rowValue);
                    }
                }

                String exportFilePath = "C:\\Users\\JSHPC\\Downloads\\개인경비 지출결의서_계정관리기술팀_정승호_2022_1월.xlsx";//companyFilePath.getText();
                XSSFWorkbook companyWorkbook = null;

                try {
                    FileInputStream fis = new FileInputStream(exportFilePath);
                    companyWorkbook = new XSSFWorkbook(fis);
                } catch (Exception ex) {
                    ex.printStackTrace();
                }

                XSSFSheet companySheet = companyWorkbook.getSheetAt(3);
                int row = 14;

                for (int i = 0; i < rowValues.size(); i++) {
                    XSSFRow curRow =  companySheet.getRow(i + row);
                    ArrayList list = (ArrayList) rowValues.get(i);
                    XSSFCell curCell = null;
                    int[] cellIdx = {0,2,3,4,5,6,7,9,10};

                    curCell = curRow.getCell(1);
                    curCell.setCellValue("라라카드");

                    for (int j = 0; j < list.size(); j++) {
                        curCell = curRow.getCell(cellIdx[j]);

                        if(cellIdx[j] == 9){
                            int obj = (int) Math.round((Double) list.get(j));
                            curCell.setCellValue(String.valueOf(obj));
                        }else{
                            curCell.setCellValue(list.get(j).toString());
                        }
                    }
                }

                try {
                    FileOutputStream outFile = new FileOutputStream(exportFilePath);
                    companyWorkbook.write(outFile);
                    outFile.close();
                } catch (IOException ex) {
                    ex.printStackTrace();
                }

                System.out.println("  ");
            }
        });
    }

    public GridBagConstraints setGbc(int wx, int wy, int gx, int gy) {
        gbc.fill = GridBagConstraints.BOTH;
        gbc.weightx = wx;
        gbc.weighty = wy;
        gbc.gridx = gx;
        gbc.gridy = gy;

        return gbc;
    }

    public static void main(String[] args) {
        new Tool();
    }
}