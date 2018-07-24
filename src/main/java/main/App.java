package main;


import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jb2011.lnf.beautyeye.BeautyEyeLNFHelper;
import org.jb2011.lnf.beautyeye.ch3_button.BEButtonUI;

import javax.swing.*;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileSystemView;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.*;
import java.util.List;

/**
 * 通过城市找到所属省份
 */
public class App extends JFrame implements ActionListener {

    private Logger logger = Logger.getLogger(App.class.getName());

    private JPanel jPanel;
    private JLabel jLabelPage;
    private JLabel jLabelCol;
    private JLabel jLabelFile;
    private JTextField jTextFieldPage;
    private JTextField jTextFieldCol;
    private JTextField jTextFieldFile;
    private JTextField jTextFieldInfo;
    private JButton jButtonSelect;
    private JButton jButtonStart;
    private JButton jButtonAdd1;   //+
    private JButton jButtonAdd2;   //+
    private JButton jButtonMinus1;   //-
    private JButton jButtonMinus2;   //-
    private JScrollPane scrollPane;
    private JTextArea jTextAreaInfo;
    private static JScrollBar scrollBar;

    public App() {
        this.setTitle("App");
        this.setSize(520, 460);
        this.setLocation(400, 200);
        this.setResizable(false);

        jPanel = new JPanel();
        jPanel.setLayout(null);

        jLabelPage = new JLabel("数据所在页");
        jLabelPage.setBounds(10, 10, 80, 30);
        jLabelPage.setFont(myFont(14));
        jPanel.add(jLabelPage);

        jButtonMinus1 = new JButton("-");
        jButtonMinus1.setBounds(90, 10, 30, 30);
        jButtonMinus1.setMargin(new Insets(0, 0, 0, 0));
        jButtonMinus1.addActionListener(this);
        jButtonMinus1.setFont(myFont(16));
        jPanel.add(jButtonMinus1);
        jTextFieldPage = new JTextField("1");
        jTextFieldPage.setBounds(120, 10, 240, 30);
        jTextFieldPage.setEditable(false);
        jTextFieldPage.setHorizontalAlignment(0);
        jPanel.add(jTextFieldPage);
        jButtonAdd1 = new JButton("+");
        jButtonAdd1.setBounds(360, 10, 30, 30);
        jButtonAdd1.setMargin(new Insets(0, 0, 0, 0));
        jButtonAdd1.setFont(myFont(16));
        jButtonAdd1.addActionListener(this);
        jPanel.add(jButtonAdd1);

        jLabelCol = new JLabel("数据所在列");
        jLabelCol.setFont(myFont(14));
        jLabelCol.setBounds(10, 55, 80, 30);
        jPanel.add(jLabelCol);

        jButtonMinus2 = new JButton("-");
        jButtonMinus2.setBounds(90, 55, 30, 30);
        jButtonMinus2.setMargin(new Insets(0, 0, 0, 0));
        jButtonMinus2.addActionListener(this);
        jButtonMinus2.setFont(myFont(16));
        jPanel.add(jButtonMinus2);
        jTextFieldCol = new JTextField("5");
        jTextFieldCol.setBounds(120, 55, 240, 30);
        jTextFieldCol.setHorizontalAlignment(0);
        jTextFieldCol.setEditable(false);
        jPanel.add(jTextFieldCol);
        jButtonAdd2 = new JButton("+");
        jButtonAdd2.setBounds(360, 55, 30, 30);
        jButtonAdd2.setMargin(new Insets(0, 0, 0, 0));
        jButtonAdd2.setFont(myFont(16));
        jButtonAdd2.addActionListener(this);
        jPanel.add(jButtonAdd2);

        jLabelFile = new JLabel("选择文件");
        jLabelFile.setFont(myFont(14));
        jLabelFile.setBounds(15, 100, 80, 30);
        jPanel.add(jLabelFile);

        jTextFieldFile = new JTextField();
        jTextFieldFile.setBounds(90, 100, 270, 30);
        jTextFieldFile.setForeground(Color.RED);
        jTextFieldFile.setEditable(false);
        jPanel.add(jTextFieldFile);

        jButtonSelect = new JButton("...");
        jButtonSelect.setBounds(360, 100, 30, 30);
        jButtonSelect.addActionListener(this);
        jPanel.add(jButtonSelect);


        jButtonStart = new JButton("Start");
        jButtonStart.setFont(myFont(20));
        jButtonStart.setBounds(90, 155, 300, 30);
        jButtonStart.addActionListener(this);
        jButtonStart.setUI(new BEButtonUI().setNormalColor(BEButtonUI.NormalColor.green));
        jPanel.add(jButtonStart);

        jTextAreaInfo = new JTextArea(440, 140);
        jTextAreaInfo.setWrapStyleWord(true); //换行方式：不分割单词
        jTextAreaInfo.setLineWrap(true); //自动换行
        jTextAreaInfo.append("就绪...\n");

        jTextFieldInfo = new JTextField();
        jTextFieldInfo.setBounds(10, 340, 440, 30);
        jTextFieldInfo.setForeground(Color.RED);
        jTextFieldInfo.setText("就绪...");
        jTextFieldInfo.setEditable(false);
        jTextFieldInfo.setFont(myFont(16));
        jPanel.add(jTextFieldInfo);

        scrollPane = new JScrollPane(jTextAreaInfo, ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS, ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
        scrollPane.setBounds(5, 200, 450, 140);
        jPanel.add(scrollPane);
        scrollBar = scrollPane.getVerticalScrollBar();

        this.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        this.setContentPane(jPanel);

    }

    private void refreshTextArea(String text) {
        jTextAreaInfo.append(text);
        jTextAreaInfo.paintImmediately(jTextAreaInfo.getBounds());
        //在更新logArea后，稍稍延时，否则getMaximum（）获得的数据可能不是最后的最大值，无法滚动到最后一行
        try {
            Thread.sleep(10);
        } catch (InterruptedException ex) {
            jTextAreaInfo.append("Exception" + ex.getMessage() + "\n");
        }
        scrollBar.setValue(scrollBar.getMaximum());
    }

    public void actionPerformed(ActionEvent e) {
        if (e.getSource() == jButtonSelect) {
            String path = selectFile();
            jTextFieldFile.setText(path);
            jTextFieldFile.setCaretPosition(0);
            jTextFieldInfo.setText("---选择Excel文件---");
        } else if (e.getSource() == jButtonAdd1) {
            controlValue(jTextFieldPage, '+');
            jTextFieldInfo.setText("---选择数据所在页---");
        } else if (e.getSource() == jButtonAdd2) {
            controlValue(jTextFieldCol, '+');
            jTextFieldInfo.setText("---选择数据所在列---");
        } else if (e.getSource() == jButtonMinus1) {
            controlValue(jTextFieldPage, '-');
            jTextFieldInfo.setText("---选择数据所在页---");
        } else if (e.getSource() == jButtonMinus2) {
            controlValue(jTextFieldCol, '-');
            jTextFieldInfo.setText("---选择数据所在列---");
        } else if (e.getSource() == jButtonStart) {
            jTextAreaInfo.setText("");
            jTextFieldInfo.setText("---执行---");
            String filePath = jTextFieldFile.getText();
            int sheetAt = Integer.valueOf(jTextFieldPage.getText());
            int columnAt = Integer.valueOf(jTextFieldCol.getText());
            //Runnable runReadExcel = () -> {
            readExcel(filePath, sheetAt, columnAt);
            //};
            //Thread thread = new Thread(runReadExcel);
            //Runnable runnable = this::printTest;
            //Thread thread1 = new Thread(runnable);
            //thread.start();
            //thread1.start();
        }
    }

    /**
     * 微调功能
     *
     * @param jTextField 要增加值的textField
     * @param c          要进行的操作(+/-)
     */
    private void controlValue(JTextField jTextField, char c) {
        int value = Integer.valueOf(jTextField.getText());
        switch (c) {
            case '+':
                value += 1;
                break;
            case '-':
                if (value == 1) break;
                else value -= 1;
            default:
                break;
        }
        jTextField.setText(String.valueOf(value));
    }

    /**
     * 选择文件
     */
    private String selectFile() {
        String path = "";
        JFileChooser fileChooser = new JFileChooser();
        FileSystemView fsv = FileSystemView.getFileSystemView();  //注意了，这里重要的一句
        //System.out.println(fsv.getHomeDirectory());                //得到桌面路径
        fileChooser.setCurrentDirectory(fsv.getHomeDirectory());
        fileChooser.setDialogTitle("请选择要处理的文件");
        fileChooser.setApproveButtonText("确定");
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        FileFilter filter = new ExtensionFileFilter("Excel格式文件", new String[]{"xls", "xlsx"});
        fileChooser.setFileFilter(filter);
        int result = fileChooser.showOpenDialog(this);
        if (JFileChooser.APPROVE_OPTION == result) {
            path = fileChooser.getSelectedFile().getPath();
            jTextAreaInfo.setText("");
            jTextAreaInfo.append("path:" + path + "\n");
        }
        return path;
    }

    private void printTest() {
        try {
            Thread.sleep(10);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        int i = 0;
        while (true) {
            i++;
            logger.info("Num:" + i);
            if (i == 12800) break;
        }
    }

    /**
     * 设置字体
     * @param fontSize 字体大小
     * @return Font
     */
    private Font myFont(int fontSize) {
        return new Font("宋体", Font.PLAIN, fontSize);
    }

    private void readExcel(String filePath, int sheetAt, int columnAt) {
        int length = filePath.split("\\.").length;
        if (length < 2) {
            jTextAreaInfo.append("Error 文件路径不正确\n");
            return;
        }

        //新建输出文件
        int dotAt = filePath.lastIndexOf(".");
        StringBuilder sb = new StringBuilder(filePath);
        sb.replace(dotAt, dotAt, "_" + Utils.getCurrentTimeStr2());
        File outFile = new File(sb.toString());

        try {
            FileInputStream fileInputStream = new FileInputStream(filePath);
            OutputStream outputStream = new FileOutputStream(outFile);
            Workbook wb = null;
            if (filePath.toLowerCase().endsWith("xls")) {
                wb = new HSSFWorkbook(fileInputStream);
            } else if (filePath.toLowerCase().endsWith("xlsx")) {
                wb = new XSSFWorkbook(fileInputStream);
            }
            Sheet sheet = wb.getSheetAt(sheetAt - 1);
            int lastRowNum = sheet.getLastRowNum();
            for (int i = 0; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                //logger.info("current row:" + i);
                Cell cell = row.getCell(columnAt - 1);
                String result = cell.getStringCellValue();
                if ("".equals(result)) continue;
                int resultLen = result.length();
                if (resultLen < 2) continue;
                else if (resultLen > 4) resultLen = 4;
                String province = "";
                Cell cellP = row.createCell(columnAt);
                for (int j = 2; j <= resultLen; j++) {
                    String city = result.substring(0, j);
                    jTextAreaInfo.append("[" + Utils.getCurrentTimeStr() + "]city:" + city + "\n");
                    jTextAreaInfo.paintImmediately(jTextAreaInfo.getBounds());
                    province = queryProvince(city);
                    if (province.length() > 0) {
                        cellP.setCellValue(province);
                        break;
                    }
                }
                if ("".equals(province)) {
                    cellP.setCellStyle(cellStyle(wb, HSSFColor.RED.index));
                } else if (province.split(";").length > 1) {
                    cellP.setCellStyle(cellStyle(wb, HSSFColor.YELLOW.index));

                }
                scrollBar.setValue(scrollBar.getMaximum());
            }

            wb.write(outputStream);
            outputStream.close();
            wb.close();
            fileInputStream.close();
            jTextAreaInfo.append("[" + Utils.getCurrentTimeStr() + "]完成...\n");
            jTextFieldInfo.setText("---完成---");
            jTextAreaInfo.setCaretPosition(jTextAreaInfo.getLineCount());


        } catch (Exception ex) {
            jTextAreaInfo.append("Exception:" + ex.getMessage() + "\n");
        }
    }

    private CellStyle cellStyle(Workbook wb, Short color) {
        CellStyle style = wb.createCellStyle();
        //颜色
        style.setFillForegroundColor(color);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setFillBackgroundColor(color);
        return style;
    }


    private String queryProvince(String city) {
        //select city_name,city_code, province_name from citys where city_name_abbr ='朝阳'
        StringBuffer provinces = new StringBuffer("");
        String sql = "select city_name,city_type, province_name from citys where is_deleted=0 and city_name_abbr =?";
        Connection conn = JDBCUtil.getAccessConnection();
        if (Objects.isNull(conn)) {
            jTextAreaInfo.append("数据连接异常，连接返回Null.\n");
            return null;
        }
        PreparedStatement statement = null;
        ResultSet resultSet = null;
        try {
            statement = conn.prepareStatement(sql);
            statement.setString(1, city);
            resultSet = statement.executeQuery();
            List<CityDTO> listCity = new ArrayList<>();
            Set<String> setProvince = new HashSet<>();
            while (resultSet.next()) {
                String cityName = resultSet.getString("city_name");
                String cityType = resultSet.getString("city_type");
                String provinceName = resultSet.getString("province_name");
                if (!setProvince.contains(provinceName)) {
                    setProvince.add(provinceName);
                    CityDTO cityDTO = new CityDTO(cityName, cityType, provinceName);
                    listCity.add(cityDTO);
                }
            }
            if (listCity.size() > 0) {
                //listCity.stream().forEach(c -> provinces.append(c.getCityName() + "[" + c.getCityType() + "]" + c.getProvinceName() + ";"));
                listCity.stream().forEach(c -> provinces.append(c.getProvinceName() + ";"));
            }
        } catch (Exception e1) {
            jTextAreaInfo.append("Exception:" + e1.getMessage() + "\n");
        } finally {
            JDBCUtil.closeConnection(conn, statement, resultSet);
        }
        return provinces.toString();
    }


    public static void main(String[] args) {
        try {
            BeautyEyeLNFHelper.launchBeautyEyeLNF();
            UIManager.put("RootPane.setupButtonVisible", false);
        } catch (Exception e) {
            e.printStackTrace();
        }
        App app = new App();
        app.setVisible(true);
    }

}
