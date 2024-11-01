package org.example;

import org.example.data.CreditDataBuilder;
import org.example.data.DepositDataBuilder;
import org.example.excel.ExcelUtils;

import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
        DepositDataBuilder depositDataBuilder = DepositDataBuilder.builder()
                .accountNumber("839128332138128")
                .firstName("Andrey")
                .lastName("Koshechkin")
                .requestId(BigDecimal.valueOf(9988133))
                .build();

        DepositDataBuilder depositDataBuilder2 = DepositDataBuilder.builder()
                .accountNumber("3333333")
                .firstName("Alex")
                .lastName("Bob")
                .requestId(BigDecimal.valueOf(1111))
                .build();
        DepositDataBuilder depositDataBuilder24 = DepositDataBuilder.builder()
                .accountNumber("3333333")
                .firstName("Alex")
                .lastName("Bob")
                .requestId(BigDecimal.valueOf(1111))
                .build();
        DepositDataBuilder depositDataBuilder44 = DepositDataBuilder.builder()
                .accountNumber("3333333")
                .firstName("Alex")
                .lastName("Bob")
                .requestId(BigDecimal.valueOf(1111))
                .build();
        DepositDataBuilder depositDataBuilder5 = DepositDataBuilder.builder()
                .accountNumber("3333333")
                .firstName("Alex")
                .lastName("Bob")
                .requestId(BigDecimal.valueOf(1111))
                .build();
        DepositDataBuilder depositDataBuilder6 = DepositDataBuilder.builder()
                .accountNumber("3333333")
                .firstName("Alex")
                .lastName("Bob")
                .requestId(BigDecimal.valueOf(1111))
                .build();
        DepositDataBuilder depositDataBuilder7 = DepositDataBuilder.builder()
                .accountNumber("3333333")
                .firstName("Alex")
                .lastName("Bob")
                .requestId(BigDecimal.valueOf(1111))
                .build();

        List<DepositDataBuilder> depositDataBuilder1 = List.of(depositDataBuilder, depositDataBuilder2, depositDataBuilder24, depositDataBuilder44, depositDataBuilder5);
        int[] columns = {1, 2, 3, 4,};

        ExcelUtils.writeDataToExcel("E:\\Новая папка", "new-test-document", depositDataBuilder1);
        // Дальше можно сохранить байты в файл или использовать их по необходимости


            List<CreditDataBuilder> creditDataList = new ArrayList<>();
            for (int i = 1; i <= 20; i++) {
                CreditDataBuilder creditData = CreditDataBuilder.builder()
                        .dateCreditOpen(LocalDate.now().minusDays(i))
                        .firstName("FirstName" + i)
                        .lastName("LastName" + i)
                        .firstAmount("1000" + i)
                        .procent("5." + i).build();
                creditDataList.add(creditData);
                ExcelUtils.writeDataToExcel("E:\\Новая папка", "new-credit", creditDataList);
        }
    }

}
