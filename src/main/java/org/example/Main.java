package org.example;

import org.example.data.DepositDataBuilder;
import org.example.excel.ExcelUtils;

import java.io.IOException;
import java.math.BigDecimal;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
        DepositDataBuilder depositDataBuilder = DepositDataBuilder.builder()
                .accountNumber("331")
                .firstName("Andrey32131231")
                .lastName("Koshechkin")
                .requestId(BigDecimal.valueOf(3123123))
                .build();
        int[] columns = {1, 2, 3, 4,};

        ExcelUtils.writeDataToExcel("E:\\Новая папка", "new-test-document", depositDataBuilder, columns);
        // Дальше можно сохранить байты в файл или использовать их по необходимости
    }
}
