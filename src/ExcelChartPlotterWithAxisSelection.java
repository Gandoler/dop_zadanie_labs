import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.data.xy.DefaultXYDataset;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelChartPlotterWithAxisSelection {

    private static JComboBox<String> xAxisComboBox;

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> createAndShowGUI());
    }

    private static void createAndShowGUI() {
        JFrame frame = new JFrame("Excel Chart Plotter");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 600);

        JFileChooser fileChooser = new JFileChooser();
        int returnValue = fileChooser.showOpenDialog(null);

        if (returnValue == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            try {
                JPanel panel = new JPanel();
                frame.getContentPane().add(panel);

                DefaultComboBoxModel<String> comboBoxModel = new DefaultComboBoxModel<>();
                xAxisComboBox = new JComboBox<>(comboBoxModel);
                panel.add(new JLabel("Choose X Axis:"));
                panel.add(xAxisComboBox);

                DefaultXYDataset dataset = createDatasetFromExcel(selectedFile);
                JFreeChart chart = ChartFactory.createXYLineChart(
                        "Excel Data Plot",
                        "X Axis",
                        "Y Axis",
                        dataset
                );
                ChartPanel chartPanel = new ChartPanel(chart);
                frame.getContentPane().add(chartPanel);

                frame.setVisible(true);

            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static DefaultXYDataset createDatasetFromExcel(File file) throws IOException {
        DefaultXYDataset dataset = new DefaultXYDataset();

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

        int selectedXAxis = xAxisComboBox.getSelectedIndex();

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
}
