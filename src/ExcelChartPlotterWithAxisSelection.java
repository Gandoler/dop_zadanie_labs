import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.data.xy.DefaultXYDataset;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelChartPlotterWithAxisSelection {

    private static JComboBox<String> xAxisComboBox;
    private static JFrame frame;
    private static Sheet sheet;
    private static File selectedFile;

    public static void main(String[] args) {
        SwingUtilities.invokeLater(ExcelChartPlotterWithAxisSelection::createAndShowGUI);
    }

    private static void createAndShowGUI() {
        frame = new JFrame("Excel Chart Plotter");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 600);

        JFileChooser fileChooser = new JFileChooser();
        int returnValue = fileChooser.showOpenDialog(null);

        if (returnValue == JFileChooser.APPROVE_OPTION) {
            selectedFile = fileChooser.getSelectedFile();
            try {
                JPanel panel = new JPanel();
                frame.getContentPane().add(panel);

                DefaultComboBoxModel<String> comboBoxModel = new DefaultComboBoxModel<>();
                xAxisComboBox = new JComboBox<>(comboBoxModel);
                panel.add(new JLabel("Choose X Axis:"));
                panel.add(xAxisComboBox);

                xAxisComboBox.addActionListener(e -> {
                    try {
                        updateChart(selectedFile);
                    } catch (IOException ex) {
                        System.out.println(ex.toString());
                    }
                });





                frame.setVisible(true);
                xAxisComboBox.removeAllItems();
                DefaultXYDataset dataset = createDatasetFromExcel(selectedFile);

                for (String columnHeader : getColumnHeadersFromDataset(dataset)) {
                    xAxisComboBox.addItem(columnHeader); // Добавление новых значений в список
                }

                initializeChart();

            } catch (IOException e) {
                System.out.println(e.toString());
            }
        }
    }
    private static List<String> getColumnHeadersFromDataset(DefaultXYDataset dataset) {
        List<String> columnHeaders = new ArrayList<>();

        for (int i = 0; i < dataset.getSeriesCount(); i++) {
            columnHeaders.add(dataset.getSeriesKey(i).toString());
        }

        return columnHeaders;
    }
    private static void initializeChart() throws IOException {
        DefaultXYDataset dataset = createDatasetFromExcel(selectedFile);

        // Получаем индекс выбранной оси X
        int selectedXAxis = xAxisComboBox.getSelectedIndex();

        JFreeChart chart = ChartFactory.createXYLineChart(
                "Excel Data Plot",
                sheet.getRow(0).getCell(selectedXAxis).getStringCellValue(), // Установка заголовка оси X
                "Y Axis",
                dataset
        );

        ChartPanel chartPanel = new ChartPanel(chart);
        frame.getContentPane().add(chartPanel);
    }
    private static void updateChart(File selectedFile) throws IOException {
        DefaultXYDataset dataset = createDatasetFromExcel(selectedFile);
        int selectedXAxis = xAxisComboBox.getSelectedIndex();

        ChartPanel existingChartPanel = (ChartPanel) frame.getContentPane().getComponent(1);
        JFreeChart chart = existingChartPanel.getChart();

        chart.getXYPlot().getDomainAxis().setLabel(sheet.getRow(0).getCell(selectedXAxis).getStringCellValue());

        chart.getXYPlot().setDataset(dataset);
    }

    private static DefaultXYDataset createDatasetFromExcel(File file) throws IOException {
        DefaultXYDataset dataset = new DefaultXYDataset();

        FileInputStream inputStream = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(inputStream);
        sheet = workbook.getSheetAt(0);

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
