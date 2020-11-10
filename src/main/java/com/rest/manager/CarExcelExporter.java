package com.rest.manager;

import com.rest.models.Car;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

public class CarExcelExporter {
    private XSSFWorkbook xssfWorkbook;
    private XSSFSheet xssfSheet;
    private List<Car> cars;

    public CarExcelExporter(List<Car> cars){
        this.cars = cars;
        xssfWorkbook = new XSSFWorkbook();
        xssfSheet = xssfWorkbook.createSheet("Cars");
    }

    private void writeHeaderRow(){
        Row row = xssfSheet.createRow(0);

        CellStyle cellStyle = xssfWorkbook.createCellStyle();
        XSSFFont font = xssfWorkbook.createFont();
        font.setBold(true);
        font.setFontHeight(14);
        cellStyle.setFont(font);

        Cell cell = row.createCell(0);
        cell.setCellValue("Car id");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(1);
        cell.setCellValue("Brand");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(2);
        cell.setCellValue("Color");
        cell.setCellStyle(cellStyle);
    }

    private void writeDataRows(){
        int rowCount = 1;
        for (Car car : cars) {
            Row row = xssfSheet.createRow(rowCount++);
            Cell cell = row.createCell(0);
            cell.setCellValue(car.getId());

            cell = row.createCell(1);
            cell.setCellValue(car.getBrand());

            cell = row.createCell(2);
            cell.setCellValue(car.getColor());
        }
    }

    public void export(HttpServletResponse response) throws IOException {
        writeHeaderRow();
        writeDataRows();

        ServletOutputStream outputStream = response.getOutputStream();
        xssfWorkbook.write(outputStream);
        xssfWorkbook.close();
        outputStream.close();
    }
}
