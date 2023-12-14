import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.renderer.xy.XYLineAndShapeRenderer;
import org.jfree.data.xy.DefaultXYDataset;

import javax.swing.*;
import java.awt.*;
import java.awt.Color;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.io.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.List;


public class ChartExample extends JFrame {
    private JComboBox<String> xAxisComboBox;
    private DefaultXYDataset dataset;
    private ChartPanel chartPanel;

    public ChartExample() {
        setTitle("PLOTTER");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        dataset = new DefaultXYDataset();
        xAxisComboBox = new JComboBox<>();
        JButton plotButton = new JButton("Plot");
        JPanel controlPanel = new JPanel();
        controlPanel.add(new JLabel("X-Axis:"));
        controlPanel.add(xAxisComboBox);
        controlPanel.add(plotButton);

        add(controlPanel, BorderLayout.NORTH);
        Boolean true_answer = false;
        do {
            try {
                JFileChooser fileChooser = new JFileChooser();
                File current_directory = new File("/Users/gl.krutoimail.ru/Desktop/bonus/");
                fileChooser.setCurrentDirectory(current_directory);
                int returnValue = fileChooser.showOpenDialog(null);
                if (!Check_file(fileChooser.getSelectedFile())) {
                    throw new Exception("Выбранный файл - неправильный");
                }
                else {
                    true_answer=true;
                }
            } catch (Exception e) {
                System.out.println(e.toString());
            }
        }while (!true_answer);


        JFileChooser fileChooser = new JFileChooser();
        File current_directory = new File("/Users/gl.krutoimail.ru/Desktop/bonus/");
        fileChooser.setCurrentDirectory(current_directory);
        int returnValue = fileChooser.showOpenDialog(null);
        plotButton.addMouseListener(new Button_listener(xAxisComboBox,fileChooser,this));


        setSize(800, 600);
        setLocationRelativeTo(null);
        setVisible(true);
    }

    public boolean Check_file(File file){
        if(!file.exists()){
            return false;
        }
        if(!(file.toString().endsWith(".txt") || file.toString().endsWith(".xlsx"))){
            return false;
        }
        return true;
    }
    private class Button_listener implements MouseListener
    {
        private JComboBox<String> xAxisComboBox;
        private  JFileChooser fileChooser;
        private ChartExample parent;
        private int Presscount = 0;
        private DefaultXYDataset xyDataset = new DefaultXYDataset();


        Button_listener(JComboBox<String> xAxisComboBox,JFileChooser fileChooser ,ChartExample parent){
            this.xAxisComboBox = xAxisComboBox;
            this.fileChooser = fileChooser;
            this.parent = parent;
        }
        @Override
        public void mouseClicked(MouseEvent e) {

        }

        @Override
        public void mousePressed(MouseEvent e) {
            if(fileChooser.getSelectedFile().toString().endsWith(".xlsx")) {
                try {
                    if (Presscount == 0) {
                        xyDataset = init(fileChooser.getSelectedFile());
                        plotChart(xyDataset);
                    } else {
                        xyDataset = reload(fileChooser.getSelectedFile());
                        plotChart_second(xyDataset);
                    }
                    xAxisComboBox.setSelectedItem(0);

                } catch (IOException ex) {
                    ex.printStackTrace();
                    JOptionPane.showMessageDialog(parent, "Ошибка при загрузке данных из файла Excel.");
                } finally {
                    parent.repaint();
                    Presscount++;
                }
            }
            else{
                try {
                    xyDataset = createDatasetFromTextFile(fileChooser.getSelectedFile());
                } catch (IOException ex) {
                    throw new RuntimeException(ex);
                }
                plotChart(xyDataset);
            }
        }

        @Override
        public void mouseReleased(MouseEvent e) {

        }

        @Override
        public void mouseEntered(MouseEvent e) {

        }

        @Override
        public void mouseExited(MouseEvent e) {

        }
    }

    private void plotChart(DefaultXYDataset defaultXYDataset) {
        JFreeChart chart = ChartFactory.createXYLineChart(
                "PLOT", // Заголовок графика
                "X-Axis", // Название оси X
                "Y-Axis", // Название оси Y
                defaultXYDataset // Датасет
        );

        if (chartPanel != null) {
            chartPanel.setChart(chart);
            chartPanel.repaint();
        } else {
            chartPanel = new ChartPanel(chart);
            add(chartPanel, BorderLayout.CENTER);
            pack();
        }

        revalidate();
        repaint();
    }

    private void plotChart_second(DefaultXYDataset defaultXYDataset) {
        plotChart(defaultXYDataset);
    }

    private DefaultXYDataset init(File file) throws IOException {
        DefaultXYDataset dataset = new DefaultXYDataset();
        JComboBox<String> xAxisComboBox = this.xAxisComboBox;
        FileInputStream inputStream = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        int rowCount = sheet.getLastRowNum() + 1;

        List<String> columnHeaders = new ArrayList<>();

        for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) {
            Cell headerCell = sheet.getRow(0).getCell(i);
            if (headerCell != null) {
                columnHeaders.add(headerCell.getStringCellValue());
            }
        }

        String[] columnHeadersArray = columnHeaders.toArray(new String[0]);
        DefaultComboBoxModel<String> comboBoxModel = new DefaultComboBoxModel<>(columnHeadersArray);
        xAxisComboBox.setModel(comboBoxModel);

        String selectedXAxisHeader = xAxisComboBox.getSelectedItem().toString();
        int selectedXAxis = -1;

        for (int i = 0; i < columnHeadersArray.length; i++) {
            if (selectedXAxisHeader.equals(columnHeadersArray[i])) {
                selectedXAxis = i;
                break;
            }
        }

        for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) {
            if (i != selectedXAxis) {
                double[][] data = new double[2][rowCount - 1]; // 2 rows (X, Y) and number of points

                for (int j = 1; j < rowCount; j++) {
                    Row row = sheet.getRow(j);
                    Cell xCell = row.getCell(selectedXAxis);
                    Cell yCell = row.getCell(i);

                    double xValue = xCell.getNumericCellValue();
                    double yValue = yCell.getNumericCellValue();

                    data[0][j - 1] = xValue;
                    data[1][j - 1] = yValue;
                }

                dataset.addSeries(sheet.getRow(0).getCell(i).getStringCellValue(), data);
            }
        }

        workbook.close();
        inputStream.close();

        return dataset;
    }

    private DefaultXYDataset reload(File file) throws IOException {
        DefaultXYDataset dataset = new DefaultXYDataset();
        JComboBox<String> xAxisComboBox = this.xAxisComboBox;
        FileInputStream inputStream = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        int rowCount = sheet.getLastRowNum() + 1;

        int newSelectedXAxis = xAxisComboBox.getSelectedIndex();

        for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) {
            if (i != newSelectedXAxis) {
                double[][] data = new double[2][rowCount - 1];

                for (int j = 1; j < rowCount; j++) {
                    Row row = sheet.getRow(j);
                    Cell xCell = row.getCell(newSelectedXAxis);
                    Cell yCell = row.getCell(i);

                    double xValue = xCell.getNumericCellValue();
                    double yValue = yCell.getNumericCellValue();

                    data[0][j - 1] = xValue;
                    data[1][j - 1] = yValue;
                }

                dataset.addSeries(sheet.getRow(0).getCell(i).getStringCellValue(), data);
            }
        }

        workbook.close();
        inputStream.close();

        return dataset;
    }
    private DefaultXYDataset createDatasetFromTextFile(File file) throws IOException {
        DefaultXYDataset dataset = new DefaultXYDataset();
        List<List<Double>> allColumns = new ArrayList<>();

        try (BufferedReader br = new BufferedReader(new FileReader(file))) {
            String line;
            while ((line = br.readLine()) != null) {
                String[] parts = line.split("\\s+");
                if (parts.length >= 2) {
                    for (int i = 0; i < parts.length; i++) {
                        if (allColumns.size() <= i) {
                            allColumns.add(new ArrayList<>());
                        }
                        allColumns.get(i).add(Double.parseDouble(parts[i]));
                    }
                }
            }
        }

        XYLineAndShapeRenderer renderer = new XYLineAndShapeRenderer();

        for (int i = 0; i < allColumns.size(); i++) {
            List<Double> columnData = allColumns.get(i);
            double[][] data = new double[2][columnData.size()];

            for (int j = 0; j < columnData.size(); j++) {
                data[0][j] = j + 1;
                data[1][j] = columnData.get(j);
            }


            Color randomColor = new Color((int) (Math.random() * 0x1000000));

            dataset.addSeries("Data from Column " + (i + 1), data);
            renderer.setSeriesPaint(i, randomColor);
        }

        return dataset;
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(ChartExample::new);
    }
}
