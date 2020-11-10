package com.rest.controllers;

import com.rest.manager.CarExcelExporter;
import com.rest.models.Car;
import com.sun.istack.internal.NotNull;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

@RestController
public class ExcelController {
    @PostMapping("/excel")
    public void downloadToExcel(@NotNull @RequestBody List<Car> cars, HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=cars.xlsx";
        response.setHeader(headerKey, headerValue);

        CarExcelExporter carExcelExporter = new CarExcelExporter(cars);
        carExcelExporter.export(response);
    }
}
